
# ============================================================
# FarmWissen-Tutorial:
# Frei verfuegbare meteorologische Daten automatisch auswerten
#
# Steuerung ueber FarmWissen_DWD_Eingabe_v4.xlsx
#
# Ausgaben je Parameter:
# 1. Jahresverlauf mit historischem Perzentilkorridor
# 2. Monatliche Darstellung mit Perzentilkorridor
# 3. Monatliche Abweichungen fuer die ausgewaehlten Jahre
# 4. Wetterkalender: eine PDF-Seite je ausgewaehltem Jahr
#
# Parameter: Niederschlag und/oder Temperatur
# Quelle: Deutscher Wetterdienst (DWD)
# ============================================================


# ---------- Pakete ----------

required_packages <- c(
  "readxl", "openxlsx" , "dplyr", "ggplot2", "lubridate", "stringr",
  "rvest", "readr", "geosphere", "tidyr", "purrr","sf","plotly"
  )

missing_packages <- required_packages[
  !vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)
]

if (length(missing_packages)) {
  install.packages(
    missing_packages,
    repos = "https://cloud.r-project.org"
  )
}

invisible(lapply(required_packages, library, character.only = TRUE))


# ---------- Dateien und Excel-Einstellungen ----------

script_dir <- getwd()

config_file <- file.path(
  script_dir,
  "FarmWissen_DWD_Eingabe.xlsx"
)

if (!file.exists(config_file)) {
  stop(
    "Die Datei 'FarmWissen_DWD_Eingabe.xlsx' wurde nicht gefunden. ",
    "Speichern Sie Excel-Datei und R-Skript im gleichen Ordner."
  )
}

cfg_raw <- readxl::read_excel(
  config_file,
  sheet = "Eingaben",
  skip = 2,
  col_names = TRUE
)

cfg <- stats::setNames(
  as.list(cfg_raw$Wert),
  cfg_raw$Eingabe
)

get_cfg <- function(name) {
  value <- cfg[[name]]
  
  if (
    is.null(value) ||
    length(value) == 0 ||
    all(is.na(value))
  ) {
    stop(
      "In der Excel-Datei fehlt ein gueltiger Wert fuer: ",
      name
    )
  }
  
  value
}

# Optionale Eingaben duerfen leer sein.
get_cfg_optional <- function(name, default = NA) {
  value <- cfg[[name]]
  
  if (
    is.null(value) ||
    length(value) == 0 ||
    all(is.na(value))
  ) {
    return(default)
    
  }
  
  value
}

as_yes_no <- function(x) {
  tolower(trimws(as.character(x))) %in%
    c("ja", "yes", "true", "1")
}


# ---------- Excel-Eingaben ----------


# Stationsname hat Vorrang vor den Koordinaten.
station_name_input <- trimws(as.character(
  get_cfg_optional("DWD-Wetterstation", "")
))

if (is.na(station_name_input)) {
  station_name_input <- ""
}

use_named_station <- nzchar(station_name_input)

# Koordinaten sind nur erforderlich, wenn kein Stationsname
# in Excel ausgewählt wurde.
lat <- suppressWarnings(
  as.numeric(get_cfg_optional("Breitengrad", NA))
)

lon <- suppressWarnings(
  as.numeric(get_cfg_optional("Laengengrad", NA))
)

do_precip <- as_yes_no(
  get_cfg("Niederschlag auswerten")
)

do_temp <- as_yes_no(
  get_cfg("Temperatur auswerten")
)

out_dir <- as.character(
  get_cfg("PDF-Ausgabeordner")
)

# Relative Ausgabeordner beziehen sich auf den Skriptordner.
if (!grepl("^(/|[A-Za-z]:)", out_dir)) {
  out_dir <- file.path(script_dir, out_dir)
}

reference_type <- tolower(trimws(
  as.character(get_cfg("Diagramm: historische Referenz"))
))

highlight_years_text <- as.character(
  get_cfg("Diagramm: hervorzuhebende Jahre")
)

highlight_years <- suppressWarnings(
  as.integer(
    trimws(
      strsplit(
        highlight_years_text,
        "+",
        fixed = TRUE
      )[[1]]
    )
  )
)

highlight_years <- unique(
  highlight_years[!is.na(highlight_years)]
)

show_labels <- as_yes_no(
  get_cfg("Diagramm: Beschriftungen")
)

precip_display <- tolower(trimws(
  as.character(get_cfg("Diagramm: Darstellung historische Referenz"))
))

if (!precip_display %in% c("korridor", "alle jahre")) {
  stop(
    "Bitte bei 'Diagramm: Darstellung historische Referenz' ",
    "entweder 'Korridor' oder 'Alle Jahre' eintragen."
  )
}

update <- as_yes_no(
  get_cfg("Aktualisierung DWD-Stationsliste und -karte")
)


# ---------- Eingaben pruefen ----------

if (!use_named_station) {
  if (!is.finite(lat) || lat < 47 || lat > 56) {
    stop(
      "Bitte einen gueltigen Breitengrad fuer Deutschland eingeben ",
      "oder in Excel eine DWD-Wetterstation auswaehlen."
    )
  }
  
  if (!is.finite(lon) || lon < 5 || lon > 16) {
    stop(
      "Bitte einen gueltigen Laengengrad fuer Deutschland eingeben ",
      "oder in Excel eine DWD-Wetterstation auswaehlen."
    )
  }
}

if (!do_precip && !do_temp) {
  stop("Waehlen Sie in Excel mindestens einen Parameter aus.")
}

if (!length(highlight_years)) {
  stop(
    "Bitte mindestens ein hervorzuhebendes Jahr eingeben, ",
    "z. B. 2024, 2025, 2026."
  )
}

if (!reference_type %in% c("median", "mittelwert")) {
  stop(
    "Bei 'Diagramm: historische Referenz' ",
    "bitte 'Median' oder 'Mittelwert' eintragen."
  )
}

if (!dir.exists(out_dir)) {
  dir.create(out_dir, recursive = TRUE)
}


# ---------- Hilfsfunktionen ----------

month_labels <- c(
  "Jan", "Feb", "Maer", "Apr", "Mai", "Jun",
  "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"
)

weekday_labels <- c("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So")

# Einheitliche historische Zusammenfassung
reference_stat <- function(x) {
  x <- x[is.finite(x)]
  
  if (!length(x)) {
    return(NA_real_)
  }
  
  if (reference_type == "median") {
    stats::median(x)
  } else {
    mean(x)
  }
}

# Quantile fuer die Korridore berechnen
quantile_summary <- function(data, group_columns, value_column) {
  data |>
    dplyr::group_by(dplyr::across(dplyr::all_of(group_columns))) |>
    dplyr::summarise(
      p10 = if (any(is.finite(.data[[value_column]]))) {
        stats::quantile(
          .data[[value_column]], 0.10,
          na.rm = TRUE, names = FALSE
        )
      } else {
        NA_real_
      },
      p25 = if (any(is.finite(.data[[value_column]]))) {
        stats::quantile(
          .data[[value_column]], 0.25,
          na.rm = TRUE, names = FALSE
        )
      } else {
        NA_real_
      },
      p50 = if (any(is.finite(.data[[value_column]]))) {
        stats::quantile(
          .data[[value_column]], 0.50,
          na.rm = TRUE, names = FALSE
        )
      } else {
        NA_real_
      },
      p75 = if (any(is.finite(.data[[value_column]]))) {
        stats::quantile(
          .data[[value_column]], 0.75,
          na.rm = TRUE, names = FALSE
        )
      } else {
        NA_real_
      },
      p90 = if (any(is.finite(.data[[value_column]]))) {
        stats::quantile(
          .data[[value_column]], 0.90,
          na.rm = TRUE, names = FALSE
        )
      } else {
        NA_real_
      },
      .groups = "drop"
    )
}

# Monatsanfang als Position auf einer gemeinsamen Nicht-Schaltjahr-Achse
month_start_positions <- as.integer(
  as.Date(paste0("2001-", sprintf("%02d", 1:12), "-01")) -
    as.Date("2001-01-01")
) + 1L


# ---------- DWD-Stationsverzeichnis laden ----------

station_meta_url <- paste0(
  "https://opendata.dwd.de/climate_environment/",
  "CDC/observations_germany/climate/daily/kl/recent/",
  "KL_Tageswerte_Beschreibung_Stationen.txt"
)

meta_file <- tempfile(fileext = ".txt")

utils::download.file(
  station_meta_url,
  meta_file,
  mode = "wb",
  quiet = TRUE
)

asd <- strsplit(
  gsub(" +", " ", readLines(meta_file)),
  " "
)

stations_tab <- lapply(asd, function(x) {
  c(
    x[c(1:6, (length(x) - sum(x %in% "Frei")))],
    paste(
      x[7:((length(x) - sum(x %in% "Frei")) - 1)],
      collapse = " "
    )
  )
})

stations <- data.frame(
  do.call(rbind, stations_tab)[-(1:2), ]
)

names(stations) <- asd[[1]][c(1:6, 8, 7)]

# ---------- Optionale Stationsliste und Stationskarte ----------

if (isTRUE(update)) {
  stations_output <- file.path(
    script_dir,
    "DWD Stationen.xlsx"
  )
  
  openxlsx::write.xlsx(
    stations |> dplyr::arrange(Stationsname),
    stations_output,
    overwrite = TRUE
  )
  
  message(
    "Aktuelle DWD-Stationsliste gespeichert: ",
    stations_output
  )
  
  shapefile_path <- file.path(
    script_dir,
    "shp",
    "NUTS250_N1.shp"
  )
  
  if (file.exists(shapefile_path)) {
    # Stationskarte nur erstellen, wenn die benoetigte Shapefile vorhanden ist.
    stations_map <- stations |>
      dplyr::mutate(
        geoBreite = suppressWarnings(as.numeric(geoBreite)),
        geoLaenge = suppressWarnings(as.numeric(geoLaenge)),
        von_date = as.Date(
          as.character(von_datum), "%Y%m%d"
        ),
        bis_date = as.Date(
          as.character(bis_datum), "%Y%m%d"
        )
      ) |>
      dplyr::filter(
        is.finite(geoBreite),
        is.finite(geoLaenge)
      )
    
    sf_wgs <- sf::st_as_sf(
      stations_map,
      coords = c("geoLaenge", "geoBreite"),
      crs = 4326,
      remove = FALSE
    )
    
    sf_wgs$Dauer_Jahre <-
      lubridate::year(sf_wgs$bis_date) -
      lubridate::year(sf_wgs$von_date)
    
    sf_wgs$Daten_aktuell <- ifelse(
      sf_wgs$bis_date > Sys.Date() - 14,
      "Ja", "Nein"
    )
    
    sf_utm <- sf::st_transform(
      sf_wgs,
      crs = 25832
    )
    
    germany_states <- sf::read_sf(shapefile_path)
    
    map_plot <- ggplot2::ggplot(
      sf_utm,
      ggplot2::aes(col = Daten_aktuell)
    ) +
      ggplot2::geom_sf(
        data = germany_states,
        fill = "grey95",
        colour = "grey60"
      ) +
      ggplot2::geom_sf(
        ggplot2::aes(
          size = Dauer_Jahre,
          text = paste0(
            "Station: ", Stationsname,
            "<br>Stations-ID: ", Stations_id,
            "<br>Breitengrad: ", geoBreite,
            "<br>Laengengrad: ", geoLaenge,
            "<br>Stationshoehe: ", Stationshoehe, " m",
            "<br>von Datum: ",
            format(von_date, "%d.%m.%Y"),
            "<br>bis Datum: ",
            format(bis_date, "%d.%m.%Y"),
            "<br>Messzeitraum: ",
            Dauer_Jahre, " Jahre"
          )
        )
      ) +
      ggplot2::labs(title = "DWD-Stationen") +
      ggplot2::scale_color_manual(
        "DWD-Stationsdaten aktuell",
        values = c(
          "Ja" = "#00BFC4",
          "Nein" = "#F8766D"
        )
      ) +
      ggplot2::scale_size(range = c(0.25, 3)) +
      ggplot2::guides(size = "none")
    
    html_plot <- plotly::ggplotly(
      map_plot,
      tooltip = "text"
    )
    
    htmlwidgets::saveWidget(
      html_plot,
      file.path(script_dir, "DWD Stationen.html"),
      selfcontained = TRUE
    )
    
    message("Interaktive DWD-Stationskarte wurde erstellt.")
  } else {
    message(
      "Hinweis: Die DWD-Stationsliste wurde aktualisiert. ",
      "Die Karte wurde nicht erstellt, weil die Datei ",
      "'shp/NUTS250_N1.shp' fehlt. Legen Sie die Shapefile ",
      "samt Begleitdateien im Unterordner 'shp' ab, ",
      "wenn Sie auch die Karte erzeugen moechten."
    )
  }
}

unlink(meta_file)


# ---------- Stationsdaten aufbereiten ----------

if (
  !all(c("Stations_id", "geoBreite", "geoLaenge") %in% names(stations)) ||
  !nrow(stations) ||
  all(is.na(stations$geoBreite))
) {
  stop(
    "Die DWD-Stationsliste konnte nicht gelesen werden. ",
    "Das Format der DWD-Datei koennte sich geaendert haben."
  )
}

stations <- stations |>
  dplyr::mutate(
    Stations_id = stringr::str_pad(
      as.character(Stations_id),
      width = 5,
      side = "left",
      pad = "0"
    ),
    geoBreite = suppressWarnings(as.numeric(geoBreite)),
    geoLaenge = suppressWarnings(as.numeric(geoLaenge)),
    von_datum = suppressWarnings(
      as.Date(as.character(von_datum), format = "%Y%m%d")
    ),
    bis_datum = suppressWarnings(
      as.Date(as.character(bis_datum), format = "%Y%m%d")
    )
  ) |>
  dplyr::filter(
    is.finite(geoBreite),
    is.finite(geoLaenge)
  )

stations <- stations |>
  dplyr::filter(bis_datum > Sys.Date() - 14)

if (!nrow(stations)) {
  stop("Keine gueltigen DWD-Stationen mit aktuellen Daten gefunden.")
}

# Stationsauswahl:
# - Stationsname vorhanden: Koordinaten ignorieren.
# - Stationsname leer: geeignete Station nach Entfernung suchen.
if (use_named_station) {
  stations <- stations |>
    dplyr::filter(
      tolower(trimws(Stationsname)) ==
        tolower(trimws(station_name_input))
    )
  
  if (!nrow(stations)) {
    stop(
      "Die in Excel ausgewaehlte DWD-Station '",
      station_name_input,
      "' ist in der aktuellen DWD-Stationsliste nicht verfuegbar. ",
      "Bitte waehlen Sie einen aktuellen Stationsnamen oder leeren Sie ",
      "das Feld 'DWD-Wetterstation', um die Stationssuche anhand ",
      "der Koordinaten zu verwenden."
    )
  }
  
  stations$distance_km <- NA_real_
  
  # Bei mehreren Eintraegen gleichen Namens den aktuellsten bevorzugen.
  stations <- stations |>
    dplyr::arrange(dplyr::desc(bis_datum))
} else {
  stations$distance_km <- geosphere::distHaversine(
    cbind(stations$geoLaenge, stations$geoBreite),
    c(lon, lat)
  ) / 1000
  
  stations <- stations |>
    dplyr::arrange(distance_km)
}


# ---------- DWD-Dateien herunterladen ----------

download_dwd_parameter <- function(
    station_id,
    parameter = c("precip", "temp")) {
  
  parameter <- match.arg(parameter)
  
  if (parameter == "precip") {
    base <- paste0(
      "https://opendata.dwd.de/climate_environment/",
      "CDC/observations_germany/climate/daily/more_precip"
    )
    prefix <- "RR"
  } else {
    base <- paste0(
      "https://opendata.dwd.de/climate_environment/",
      "CDC/observations_germany/climate/daily/kl"
    )
    prefix <- "KL"
  }
  
  fetch_dir <- function(period) {
    
    url <- paste0(base, "/", period, "/")
    
    page <- tryCatch(
      rvest::read_html(url),
      error = function(e) NULL
    )
    
    if (is.null(page)) {
      return(NULL)
    }
    
    links <- rvest::html_elements(page, "a") |>
      rvest::html_attr("href")
    
    links <- links[
      !is.na(links) &
        grepl("\\.zip$", links, ignore.case = TRUE)
    ]
    
    hit <- links[
      grepl(
        paste0(prefix, "_", station_id, "_"),
        links,
        fixed = TRUE
      )
    ]
    
    if (!length(hit)) {
      return(NULL)
    }
    
    zip_url <- paste0(url, hit[length(hit)])
    z <- tempfile(fileext = ".zip")
    
    ok <- tryCatch({
      utils::download.file(
        zip_url,
        z,
        mode = "wb",
        quiet = TRUE
      )
      TRUE
    }, error = function(e) FALSE)
    
    if (!ok || !file.exists(z) || file.info(z)$size == 0) {
      unlink(z)
      return(NULL)
    }
    
    contents <- tryCatch(
      utils::unzip(z, list = TRUE)$Name,
      error = function(e) character()
    )
    
    data_member <- contents[
      grepl("produkt.*\\.txt$", contents, ignore.case = TRUE)
    ]
    
    if (!length(data_member)) {
      unlink(z)
      return(NULL)
    }
    
    dat <- tryCatch(
      readr::read_delim(
        unz(z, data_member[1]),
        delim = ";",
        show_col_types = FALSE,
        na = c("-999", "-999.0", "-999.00", "")
      ),
      error = function(e) NULL
    )
    
    unlink(z)
    dat
  }
  
  dat <- list(
    fetch_dir("historical"),
    fetch_dir("recent")
  )
  
  dat <- dat[!vapply(dat, is.null, logical(1))]
  
  if (!length(dat)) {
    return(NULL)
  }
  
  dplyr::bind_rows(dat) |>
    dplyr::distinct()
}


# ---------- Daten vorbereiten ----------

make_plot_data <- function(raw, parameter, selected_station) {
  
  if (is.null(raw) || !nrow(raw)) {
    stop(
      "Fuer Station ", selected_station$Stations_id,
      " sind keine ", parameter, "-Daten abrufbar."
    )
  }
  
  if (!"MESS_DATUM" %in% names(raw)) {
    stop("In den heruntergeladenen DWD-Daten fehlt MESS_DATUM.")
  }
  
  if (parameter == "precip") {
    
    value_col <- if ("RS" %in% names(raw)) {
      "RS"
    } else if ("RWS" %in% names(raw)) {
      "RWS"
    } else {
      NA_character_
    }
    
    if (is.na(value_col)) {
      stop(
        "In den Niederschlagsdaten wurde keine passende ",
        "Messwertspalte (RS/RWS) gefunden."
      )
    }
    
    cumulative_type <- "sum"
    
  } else {
    
    value_col <- if ("TMK" %in% names(raw)) {
      "TMK"
    } else {
      NA_character_
    }
    
    if (is.na(value_col)) {
      stop("In den Temperaturdaten wurde die Spalte TMK nicht gefunden.")
    }
    
    cumulative_type <- "mean"
  }
  
  all_data <- raw |>
    dplyr::mutate(
      date = as.Date(
        as.character(MESS_DATUM),
        format = "%Y%m%d"
      ),
      value = suppressWarnings(as.numeric(.data[[value_col]])),
      value = dplyr::if_else(
        !is.finite(value) | value <= -999,
        NA_real_,
        value
      ),
      Year = lubridate::year(date),
      DoY = lubridate::yday(date),
      Month = lubridate::month(date),
      MonthName = factor(
        Month,
        levels = 1:12,
        labels = month_labels
      ),
      DoM = lubridate::day(date),
      
      # Gemeinsame Nicht-Schaltjahr-Achse fuer Jahresdiagramme.
      # Der 29. Februar wird auf den 28. Februar gelegt.
      PlotDate = as.Date(
        ifelse(
          format(date, "%m-%d") == "02-29",
          paste0(Year, "-02-28"),
          paste0("2001-", format(date, "%m-%d"))
        ),
        format = "%Y-%m-%d"
      ),
      PlotDay = as.integer(PlotDate - as.Date("2001-01-01")) + 1L
    ) |>
    dplyr::filter(!is.na(date)) |>
    dplyr::arrange(date) |>
    dplyr::distinct(date, .keep_all = TRUE)
  
  current_year <- lubridate::year(Sys.Date())
  
  # Nur abgeschlossene Jahre mit allen Tagesmesswerten als Referenz
  completeness <- all_data |>
    dplyr::filter(!is.na(value), Year < current_year) |>
    dplyr::group_by(Year) |>
    dplyr::summarise(
      valid_days = dplyr::n_distinct(date),
      expected_days = ifelse(
        lubridate::leap_year(dplyr::first(Year)),
        366L,
        365L
      ),
      .groups = "drop"
    ) |>
    dplyr::filter(valid_days == expected_days) |>
    dplyr::pull(Year)
  
  usable_years <- c(completeness, current_year)
  
  all_data <- all_data |>
    dplyr::filter(Year %in% usable_years) |>
    dplyr::group_by(Year) |>
    dplyr::arrange(date, .by_group = TRUE) |>
    dplyr::mutate(
      plotted_value = if (cumulative_type == "sum") {
        cumsum(dplyr::coalesce(value, 0))
      } else {
        cumsum(dplyr::coalesce(value, 0)) /
          cumsum(!is.na(value))
      }
    ) |>
    dplyr::ungroup() |>
    dplyr::group_by(Year, Month) |>
    dplyr::arrange(date, .by_group = TRUE) |>
    dplyr::mutate(
      monthly_value = if (cumulative_type == "sum") {
        cumsum(dplyr::coalesce(value, 0))
      } else {
        cumsum(dplyr::coalesce(value, 0)) /
          cumsum(!is.na(value))
      }
    ) |>
    dplyr::ungroup()
  
  if (!nrow(all_data)) {
    stop(
      "Nach der Pruefung auf vollstaendige Jahre ",
      "bleiben keine Daten uebrig."
    )
  }
  
  attr(all_data, "complete_years") <- sort(unique(completeness))
  attr(all_data, "cumulative_type") <- cumulative_type
  
  all_data
}


# ---------- Tages- und Monatsdiagramme ----------

plot_time_series <- function(dat, parameter, selected_station) {
  
  if (!nrow(dat)) {
    stop("Keine Messwerte fuer die Darstellung vorhanden.")
  }
  
  years <- sort(unique(dat$Year))
  current_year <- lubridate::year(Sys.Date())
  highlighted <- intersect(highlight_years, years)
  
  if (!length(highlighted)) {
    warning(
      "Keines der in Excel hervorgehobenen Jahre ",
      "ist in den auswertbaren Daten enthalten."
    )
  }
  
  # Hervorgehobene Jahre und aktuelles Jahr sind keine Referenzjahre.
  reference_years <- sort(
    setdiff(years, c(highlighted, current_year))
  )
  
  ref_label <- if (reference_type == "median") {
    "Historischer Median"
  } else {
    "Historischer Mittelwert"
  }
  
  reference_period <- if (length(reference_years)) {
    paste0(
      min(reference_years), "-", max(reference_years),
      " (", length(reference_years), " vollstaendige Jahre)"
    )
  } else {
    "keine Referenzjahre"
  }
  
  # Excel-Einstellung fuer Niederschlag; fuer Temperatur ebenfalls
  # verwendbar, damit die Darstellung einheitlich bleibt.
  display_all_years <- precip_display == "alle jahre"

  
  y_label <- if (parameter == "precip") {
    "Kumulierter Niederschlag seit Jahresbeginn [mm]"
  } else {
    "Mittlere Temperatur seit Jahresbeginn [°C]"
  }
  
  monthly_y_label <- if (parameter == "precip") {
    "Kumulierter Niederschlag seit Monatsbeginn [mm]"
  } else {
    "Mittlere Temperatur seit Monatsbeginn [°C]"
  }
  
  title_text <- if (parameter == "precip") {
    "Niederschlag im historischen Vergleich"
  } else {
    "Temperatur im historischen Vergleich"
  }
  
  station_subtitle <- 
    ifelse(use_named_station,
           paste0(
             selected_station$Stationsname,
             " (DWD ", selected_station$Stations_id,")"),
    paste0(
    selected_station$Stationsname,
    " (DWD ", selected_station$Stations_id, "; ",
    round(selected_station$distance_km, 1), " km entfernt)"
  ))
  
  ifelse(display_all_years,
    caption_text <- paste0(
      "Referenz: ", reference_period,
      ". Referenzlinie: ", ref_label,
      ". Einzeljahre.\nDaten-Quelle: DWD. Für die Richtigkeit, Vollständigkeit und Aktualität der visuellen Darstellung wird keine Gewähr übernommen."
    ),
  caption_text <- paste0(
    "Referenz: ", reference_period,
    ". Referenzlinie: ", ref_label,
    ". Korridore: P25-P75 (dunkelgrau) und P10-P90 (hellgrau).\nDaten-Quelle: DWD. Für die Richtigkeit, Vollständigkeit und Aktualität der visuellen Darstellung wird keine Gewähr übernommen."
  ))
  
  
  # ---------- Historische Jahresreferenz und Perzentile ----------
  
  daily_history <- dat |>
    dplyr::filter(Year %in% reference_years)
  
  daily_ref <- daily_history |>
    dplyr::group_by(PlotDay) |>
    dplyr::summarise(
      ref_value = reference_stat(plotted_value),
      .groups = "drop"
    )
  
  daily_corridor <- quantile_summary(
    daily_history,
    "PlotDay",
    "plotted_value"
  )
  
  
  # ---------- Plot 1: Jahresverlauf ----------
  
  p_daily <- ggplot2::ggplot()
  
  if (nrow(daily_corridor) > 0&!display_all_years) {
    p_daily <- p_daily +
      ggplot2::geom_ribbon(
        data = daily_corridor,
        ggplot2::aes(
          x = PlotDay,
          ymin = p10,
          ymax = p90
        ),
        fill = "lightgrey",
        alpha = 0.55,
        na.rm = TRUE
      ) +
      ggplot2::geom_ribbon(
        data = daily_corridor,
        ggplot2::aes(
          x = PlotDay,
          ymin = p25,
          ymax = p75
        ),
        fill = "grey",
        alpha = 0.35,
        na.rm = TRUE
      ) #+
      # ggplot2::geom_line(
      #   data = daily_corridor,
      #   ggplot2::aes(x = PlotDay, y = p50),
      #   colour = "#08519C",
      #   linewidth = 0.85,
      #   na.rm = TRUE
      # )
  }
  
  if (display_all_years && length(reference_years)) {
    p_daily <- p_daily +
      ggplot2::geom_line(
        data = daily_history,
        ggplot2::aes(
          x = PlotDay,
          y = plotted_value,
          group = factor(Year)
        ),
        colour = "grey55",
        alpha = 0.55,
        linewidth = 0.35,
        na.rm = TRUE
      )
  }
  
  if (nrow(daily_ref) > 0) {
    p_daily <- p_daily +
      ggplot2::geom_line(
        data = daily_ref,
        ggplot2::aes(x = PlotDay, y = ref_value),
        colour = "firebrick",
        linewidth = 0.9,
        linetype = "dashed",
        na.rm = TRUE
      )
  }
  
  highlighted_data <- dat |>
    dplyr::filter(Year %in% highlighted)
  
  p_daily <- p_daily +
    ggplot2::geom_line(
      data = highlighted_data,
      ggplot2::aes(
        x = PlotDay,
        y = plotted_value,
        colour = factor(Year),
        group = factor(Year)
      ),
      linewidth = 0.9,
      na.rm = TRUE
    ) +
    ggplot2::scale_x_continuous(
      breaks = month_start_positions,
      labels = paste0("1.",1:12, "."),
      limits = c(1, 365),
      expand = ggplot2::expansion(mult = c(0.005, 0.10))
    ) +
    ggplot2::labs(
      title = paste0(title_text, " - jährliche Darstellung"),
      subtitle = station_subtitle,
      x = "Monat",
      y = y_label,
      colour = "Hervorgehobenes Jahr",
      caption = caption_text
    ) +
    ggplot2::coord_cartesian(clip = "off") +
    ggplot2::theme_minimal(base_size = 11) +
    ggplot2::theme(
      legend.position = "bottom",
      plot.caption = ggplot2::element_text(size = 8),
      plot.margin = ggplot2::margin(8, 35, 8, 8)
    )
  
  if (show_labels && nrow(highlighted_data) > 0) {
    
    last_values <- highlighted_data |>
      dplyr::filter(!is.na(plotted_value)) |>
      dplyr::group_by(Year) |>
      dplyr::filter(date == max(date)) |>
      dplyr::ungroup()
    
    # # attach average value
    # last_values <- bind_rows(last_values %>% mutate(Year=as.character(Year)),daily_ref %>% filter(PlotDay%in%c(365,366)) %>% 
    #                                     filter(PlotDay==max(PlotDay)) %>% 
    #                                     mutate(plotted_value=ref_value,
    #                                            Year="Ref"))
    
    p_daily <- p_daily +
      ggplot2::geom_label(
        data = last_values,
        ggplot2::aes(
          x = PlotDay,
          y = plotted_value,
          label = paste0(Year, ": ", round(plotted_value, 1)),
          colour = factor(Year)
        ),
        hjust = -0.05,
        size = 2.8,
        show.legend = FALSE,
        label.size = 0
      )
  }
  
  
  # ---------- Plot 2: Monatsverlauf ----------
  
  monthly_history <- dat |>
    dplyr::filter(
      Year %in% reference_years,
      !(DoM == 29 & Month == 2)
    )
  
  monthly_ref <- monthly_history |>
    dplyr::group_by(Month, DoM, MonthName) |>
    dplyr::summarise(
      ref_value = reference_stat(monthly_value),
      .groups = "drop"
    )
  
  monthly_corridor <- quantile_summary(
    monthly_history,
    c("Month", "DoM", "MonthName"),
    "monthly_value"
  )
  
  p_monthly <- ggplot2::ggplot()
  
  if (nrow(monthly_corridor) > 0 & !display_all_years) {
    p_monthly <- p_monthly +
      ggplot2::geom_ribbon(
        data = monthly_corridor,
        ggplot2::aes(
          x = DoM,
          ymin = p10,
          ymax = p90,
          group = Month
        ),
        fill = "lightgrey",
        alpha = 0.55,
        na.rm = TRUE
      ) +
      ggplot2::geom_ribbon(
        data = monthly_corridor,
        ggplot2::aes(
          x = DoM,
          ymin = p25,
          ymax = p75,
          group = Month
        ),
        fill = "grey",
        alpha = 0.35,
        na.rm = TRUE
      ) #+
      # ggplot2::geom_line(
      #   data = monthly_corridor,
      #   ggplot2::aes(
      #     x = DoM,
      #     y = p50,
      #     group = Month
      #   ),
      #   colour = "#08519C",
      #   linewidth = 0.75,
      #   na.rm = TRUE
      # )
  }
  
  if (display_all_years && length(reference_years)) {
    p_monthly <- p_monthly +
      ggplot2::geom_line(
        data = monthly_history,
        ggplot2::aes(
          x = DoM,
          y = monthly_value,
          group = interaction(Year, Month)
        ),
        colour = "grey55",
        alpha = 0.45,
        linewidth = 0.3,
        na.rm = TRUE
      )
  }
  
  if (nrow(monthly_ref) > 0) {
    p_monthly <- p_monthly +
      ggplot2::geom_line(
        data = monthly_ref,
        ggplot2::aes(x = DoM, y = ref_value, group = Month),
        colour = "firebrick",
        linewidth = 0.8,
        linetype = "dashed",
        na.rm = TRUE
      )
  }
  
  p_monthly <- p_monthly +
    ggplot2::geom_line(
      data = highlighted_data,
      ggplot2::aes(
        x = DoM,
        y = monthly_value,
        colour = factor(Year),
        group = interaction(Year, Month)
      ),
      linewidth = 0.7,
      na.rm = TRUE
    ) +
    ggplot2::facet_wrap(
      ggplot2::vars(MonthName),
      nrow = 1,
      scales = "free_x",
      drop = FALSE
    ) +
    ggplot2::labs(
      title = paste0(title_text, " - monatliche Darstellung"),
      subtitle = station_subtitle,
      x = "Tag des Monats",
      y = monthly_y_label,
      colour = "Hervorgehobenes Jahr",
      caption = caption_text
    ) +
    ggplot2::theme_minimal(base_size = 10) +
    ggplot2::theme(
      legend.position = "bottom",
      strip.text = ggplot2::element_text(face = "bold"),
      plot.caption = ggplot2::element_text(size = 7),
      axis.text.x = ggplot2::element_text(size = 7),
      panel.border = ggplot2::element_rect(
        colour = "grey85",
        fill = NA,
        linewidth = 0.3
      )
    )
  
  list(
    daily = p_daily,
    monthly = p_monthly,
    reference_years = reference_years,
    reference_period = reference_period
  )
}


# ---------- Monatliche Abweichungen fuer ausgewaehlte Jahre ----------

plot_monthly_anomaly <- function(dat, parameter, selected_station) {
  
  current_year <- lubridate::year(Sys.Date())
  years <- sort(unique(dat$Year))
  
  selected_years <- intersect(highlight_years, years)
  
  reference_years <- setdiff(
    years,
    c(selected_years, current_year)
  )
  
  title_text <- if (parameter == "precip") {
    "Monatliche Niederschlagsabweichung"
  } else {
    "Monatliche Temperaturabweichung"
  }
  
  station_subtitle <- 
    ifelse(use_named_station,
           paste0(
             selected_station$Stationsname,
             " (DWD ", selected_station$Stations_id,")"),
           paste0(
             selected_station$Stationsname,
             " (DWD ", selected_station$Stations_id, "; ",
             round(selected_station$distance_km, 1), " km entfernt)"
           ))
  
  empty_plot <- function(message_text) {
    ggplot2::ggplot() +
      ggplot2::annotate(
        "text",
        x = 1,
        y = 1,
        label = message_text,
        size = 4
      ) +
      ggplot2::labs(
        title = title_text,
        subtitle = station_subtitle
      ) +
      ggplot2::theme_void()
  }
  
  if (!length(selected_years)) {
    return(empty_plot("Keine Messwerte fuer die ausgewaehlten Jahre verfuegbar."))
  }
  
  if (!length(reference_years)) {
    return(empty_plot("Keine historischen Referenzjahre verfuegbar."))
  }
  
  selected_data <- dat |>
    dplyr::filter(
      Year %in% selected_years,
      !is.na(value)
    )
  
  # Letzter verfuegbarer Messtag je Monat und ausgewaehltem Jahr
  cutoffs <- selected_data |>
    dplyr::group_by(Year, Month, MonthName) |>
    dplyr::summarise(
      cutoff_day = max(DoM),
      .groups = "drop"
    ) |>
    dplyr::rename(target_year = Year)
  
  # Werte des ausgewaehlten Jahres bis zum jeweiligen Monatsstichtag
  selected_monthly <- selected_data |>
    dplyr::rename(target_year = Year) |>
    dplyr::inner_join(
      cutoffs,
      by = c("target_year", "Month", "MonthName")
    ) |>
    dplyr::filter(DoM <= cutoff_day) |>
    dplyr::group_by(target_year, Month, MonthName, cutoff_day) |>
    dplyr::summarise(
      selected_value = if (parameter == "precip") {
        sum(value, na.rm = TRUE)
      } else {
        mean(value, na.rm = TRUE)
      },
      .groups = "drop"
    )
  
  # Historische Monatswerte jeweils bis zum gleichen Kalendertag
  historical_monthly <- dat |> 
    dplyr::filter(
      Year %in% reference_years,
      !is.na(value)
    ) |>
    dplyr::rename(reference_year = Year) |>
    dplyr::inner_join(
      cutoffs%>% dplyr::select(cutoff_day,Month,target_year),
      by = "Month"
    ) |>
    dplyr::filter(DoM <= cutoff_day) |>
    dplyr::group_by(
      target_year, Month, MonthName,
      cutoff_day, reference_year
    ) |>
    dplyr::summarise(
      historical_value = if (parameter == "precip") {
        sum(value, na.rm = TRUE)
      } else {
        mean(value, na.rm = TRUE)
      },
      .groups = "drop"
    )
  
  historical_reference <- historical_monthly |>
    dplyr::group_by(target_year, Month, MonthName, cutoff_day) |>
    dplyr::summarise(
      reference_value = reference_stat(historical_value),
      .groups = "drop"
    )
  
  anomaly_data <- selected_monthly |>
    dplyr::left_join(
      historical_reference,
      by = c("target_year", "Month", "MonthName", "cutoff_day")
    ) |>
    dplyr::mutate(
      anomaly = selected_value - reference_value,
      direction = ifelse(anomaly >= 0, "Positiv", "Negativ"),
      Year = factor(target_year, levels = selected_years)
    ) |>
    dplyr::filter(is.finite(anomaly))
  
  if (!nrow(anomaly_data)) {
    return(empty_plot("Keine vergleichbaren Monatswerte verfuegbar."))
  }
  
  unit_label <- if (parameter == "precip") {
    "Abweichung Niederschlag [mm]"
  } else {
    "Abweichung Temperatur [°C]"
  }
  
  ggplot2::ggplot(
    anomaly_data,
    ggplot2::aes(
      x = MonthName,
      y = anomaly,
      fill = direction
    )
  ) +
    ggplot2::geom_col(show.legend = TRUE) +
    ggplot2::geom_hline(
      yintercept = 0,
      colour = "grey30",
      linewidth = 0.5
    ) +
    ggplot2::facet_wrap(
      ggplot2::vars(Year),
      ncol = 1
    ) +
    ggplot2::scale_fill_manual(
      values = c(
        "Positiv" = ifelse(parameter == "precip","steelblue","firebrick"),
        "Negativ" = ifelse(parameter == "precip","firebrick","steelblue")
      ),
      breaks = c("Positiv", "Negativ"),
      name = "Abweichung"
    ) +
    ggplot2::labs(
      title = title_text,
      subtitle = station_subtitle,
      x = "Monat",
      y = unit_label,
      caption = paste(
        "Positive Abweichungen blau, negative Abweichungen rot.",
        paste0("Referenz: ", paste0(
          min(historical_monthly$reference_year), "-", max(historical_monthly$reference_year),
          " (", length(unique(historical_monthly$reference_year)), " vollstaendige Jahre)"
        )),
        "des jeweiligen Monats.\nDaten-Quelle: DWD. Für die Richtigkeit, Vollständigkeit und Aktualität der visuellen Darstellung wird keine Gewähr übernommen."
      )
    ) +
    ggplot2::theme_minimal(base_size = 11) +
    ggplot2::theme(
      legend.position = "bottom",
      axis.text.x = ggplot2::element_text(
        angle = 45,
        hjust = 1
      ),
      strip.text = ggplot2::element_text(face = "bold"),
      plot.caption = ggplot2::element_text(size = 8)
    )
}


# ---------- Wetterkalender: eine Funktion je Jahr ----------

plot_weather_heatmap <- function(dat, parameter, selected_station) {
  
  years_to_plot <- intersect(
    highlight_years,
    sort(unique(dat$Year))
  )
  
  # Gemeinsame Niederschlagsskala fuer alle ausgewaehlten Jahre
  calendar_data <- dat |>
    dplyr::filter(Year %in% years_to_plot)
  
  precip_max <- suppressWarnings(
    max(calendar_data$value, na.rm = TRUE)
  )
  
  if (!is.finite(precip_max) || precip_max <= 0) {
    precip_max <- 1
  }
  
  parameter_label <- if (parameter == "precip") {
    "Täglicher Niederschlag"
  } else {
    "Tägliche Mitteltemperatur"
  }
  
  unit_label <- if (parameter == "precip") {
    "Niederschlag [mm]"
  } else {
    "Temperatur [°C]"
  }
  
  station_subtitle <- 
    ifelse(use_named_station,
           paste0(
             selected_station$Stationsname,
             " (DWD ", selected_station$Stations_id,")"),
           paste0(
             selected_station$Stationsname,
             " (DWD ", selected_station$Stations_id, "; ",
             round(selected_station$distance_km, 1), " km entfernt)"
           ))
  
  if (!length(years_to_plot)) {
    return(
      list(
        keine_daten = ggplot2::ggplot() +
          ggplot2::annotate(
            "text",
            x = 1,
            y = 1,
            label = "Keine Messwerte fuer die ausgewaehlten Jahre verfuegbar."
          ) +
          ggplot2::theme_void()
      )
    )
  }
  
  
  # Gemeinsame Farbskalen fuer alle ausgewaehlten Kalenderjahre
  calendar_data <- dat |>
    dplyr::filter(Year %in% years_to_plot)
  
  # Niederschlag: gleicher Maximalwert fuer alle Jahreskalender
  precip_max <- suppressWarnings(
    max(calendar_data$value, na.rm = TRUE)
  )
  
  if (!is.finite(precip_max) || precip_max <= 0) {
    precip_max <- 1
  }
  
  # Auf einen gut lesbaren oberen Skalenwert aufrunden
  precip_max <- ceiling(precip_max / 10) * 10
  
  # Temperatur: feste Skala fuer alle Jahre
  # Werte ausserhalb des Bereichs werden auf die Randfarbe begrenzt.
  temp_limits <- c(-20, 40)
  
  
  plots_by_year <- list()
  
  for (year_to_plot in years_to_plot) {
    
    year_data <- dat |>
      dplyr::filter(Year == year_to_plot) |>
      dplyr::mutate(
        weekday_number = lubridate::wday(
          date,
          week_start = 1
        ),
        first_weekday = lubridate::wday(
          lubridate::floor_date(date, "month"),
          week_start = 1
        ),
        calendar_week = ceiling(
          (DoM + first_weekday - 1) / 7
        ),
        # Zeilenposition 1 ist die oberste Kalenderwoche
        calendar_row = 7 - calendar_week,
        weekday_label = factor(
          weekday_labels[weekday_number],
          levels = rev(weekday_labels)
        ),
        MonthName = factor(
          Month,
          levels = 1:12,
          labels = month_labels
        ),
        day_label = as.character(DoM)
      )
    
    p <- ggplot2::ggplot(
      year_data,
      ggplot2::aes(
        x = weekday_number,
        y = calendar_row,
        fill = value
      )
    ) +
      ggplot2::geom_tile(
        colour = "white",
        linewidth = 0.45,
        width = 0.96,
        height = 0.96
      ) +
      ggplot2::geom_text(
        ggplot2::aes(label = day_label),
        size = 2.5,
        na.rm = TRUE
      ) +
      ggplot2::facet_wrap(
        ggplot2::vars(MonthName),
        ncol = 4,
        nrow = 3,
        drop = FALSE
      ) +
      ggplot2::scale_x_continuous(
        breaks = 1:7,
        labels = weekday_labels,
        limits = c(0.5, 7.5),
        expand = c(0, 0)
      ) +
      ggplot2::scale_y_continuous(
        breaks = 1:6,
        labels = 6:1,
        limits = c(0.5, 6.5),
        expand = c(0, 0)
      ) +
      ggplot2::labs(
        title = paste0(parameter_label, " - Kalender ", year_to_plot),
        subtitle = station_subtitle,
        x = NULL,
        y = NULL,
        fill = unit_label,
        caption = paste(
          "Graue Felder kennzeichnen fehlende Messwerte.",
          "\nDaten-Quelle: DWD. Für die Richtigkeit, Vollständigkeit und Aktualität der visuellen Darstellung wird keine Gewähr übernommen."
        )
      ) +
      ggplot2::theme_minimal(base_size = 10) +
      ggplot2::theme(
        strip.text = ggplot2::element_text(face = "bold", size = 10),
        axis.text.x = ggplot2::element_text(size = 7),
        axis.text.y = ggplot2::element_blank(),
        axis.ticks.y = ggplot2::element_blank(),
        panel.grid = ggplot2::element_blank(),
        panel.border = ggplot2::element_rect(
          colour = "grey45",
          fill = NA,
          linewidth = 0.5
        ),
        panel.spacing = grid::unit(0.25, "cm"),
        legend.position = "right",
        plot.caption = ggplot2::element_text(size = 8)
      )
    
    if (parameter == "precip") {
      
      p <- p +
        ggplot2::scale_fill_gradientn(
          colours = c(
            "#FFFFFF",  # 0 mm: weiß
            "#67A9CF",  # wenig Niederschlag: hellblau
            "#2166AC",  # mittlerer Niederschlag: blau
            "#B2182B",  # hoher Niederschlag: rot
            "#8E44AD"   # sehr hoher Niederschlag: lila
          ),
          limits = c(0, precip_max),
          breaks = pretty(c(0, precip_max), n = 5),
          oob = scales::squish,
          na.value = "grey85",
          name = "Niederschlag [mm]"
        )
      
    } else {
      
      p <- p +
        ggplot2::scale_fill_gradientn(
          colours = c(
            "#2166AC",  # -20 °C: dunkelblau
            "#67A9CF",  # -10 °C: hellblau
            "#FFFFFF",  #   0 °C: weiß
            "#ffff00",  #  10 °C: gelb
            "#FDAE61",  #  20 °C: orange
            "#B2182B",  #  30 °C: dunkelrot
            "#EE82EE"   #  40 °C: violet
          ),
          values = scales::rescale(
            c(-20,-10, 0, 10, 20, 30, 40),
            from = temp_limits
          ),
          limits = temp_limits,
          breaks = c(-20,-10, 0, 10, 20, 30, 40),
          labels = c("-20 °C","-10 °C", "0 °C", "10 °C", "20 °C", "30 °C","40 °C"),
          oob = scales::squish,
          na.value = "grey85",
          name = "Temperatur"
        )
    }
    
    plots_by_year[[as.character(year_to_plot)]] <- p
  }
  
  # Wichtig: eine Liste einzelner Jahresplots zur PDF-Ausgabe
  plots_by_year
}


# ---------- Station und Auswertung ----------

selected_station <- NULL
raw_precip <- NULL
raw_temp <- NULL

for (i in seq_len(nrow(stations))) {
  
  candidate <- stations[i, , drop = FALSE]
  candidate_data_ok <- TRUE
  
  if (do_precip) {
    test <- tryCatch(
      download_dwd_parameter(candidate$Stations_id, "precip"),
      error = function(e) NULL
    )
    
    if (!is.null(test) && nrow(test)) {
      raw_precip <- test
    } else {
      candidate_data_ok <- FALSE
    }
  }
  
  if (do_temp) {
    test <- tryCatch(
      download_dwd_parameter(candidate$Stations_id, "temp"),
      error = function(e) NULL
    )
    
    if (!is.null(test) && nrow(test)) {
      raw_temp <- test
    } else {
      candidate_data_ok <- FALSE
    }
  }
  
  if (
    candidate_data_ok &&
    (!do_precip || !is.null(raw_precip)) &&
    (!do_temp || !is.null(raw_temp))
  ) {
    selected_station <- candidate
    break
  } else {
    # Keine Daten unterschiedlicher Stationen kombinieren
    raw_precip <- NULL
    raw_temp <- NULL
  }
}

if (is.null(selected_station)) {
  if (use_named_station) {
    stop(
      "Fuer die ausgewaehlte DWD-Station '",
      station_name_input,
      "' konnten nicht alle ausgewaehlten Parameter abgerufen werden. ",
      "Waehlen Sie eine andere Station oder leeren Sie das Stationsfeld, ",
      "um die Stationssuche anhand der Koordinaten zu verwenden."
    )
  } else {
    stop(
      "Es wurde keine nahegelegene DWD-Station gefunden, ",
      "fuer die alle ausgewaehlten Parameter abrufbar sind."
    )
  }
}

if (do_precip && !is.null(raw_precip)) {
  names(raw_precip) <- gsub(" ", "", names(raw_precip))
}

if (do_temp && !is.null(raw_temp)) {
  names(raw_temp) <- gsub(" ", "", names(raw_temp))
}

if (is.finite(selected_station$distance_km)) {
  message(
    "Verwendete DWD-Station: ",
    selected_station$Stationsname,
    " (", selected_station$Stations_id,
    "), Entfernung ca. ",
    round(selected_station$distance_km, 1),
    " km."
  )
} else {
  message(
    "Verwendete DWD-Station: ",
    selected_station$Stationsname,
    " (", selected_station$Stations_id,
    "); Stationsname wurde in Excel direkt ausgewaehlt. ",
    "Die Koordinaten wurden ignoriert."
  )
}
# ---------- PDF-Ausgabe ----------

date_tag <- format(Sys.Date(), "%Y-%m-%d")

save_weather_pdf <- function(filename, plot_list) {
  
  grDevices::pdf(
    filename,
    width = 12,
    height = 7.5,
    onefile = TRUE
  )
  
  on.exit(grDevices::dev.off(), add = TRUE)
  
  # Feste Reihenfolge der Zeitreihen
  plot_order <- c("daily", "monthly", "anomaly")
  
  for (plot_name in plot_order) {
    
    current_plot <- plot_list[[plot_name]]
    
    if (
      !is.null(current_plot) &&
      inherits(current_plot, "ggplot")
    ) {
      print(current_plot)
    }
  }
  
  # Jeder Wetterkalender wird separat auf einer eigenen PDF-Seite gedruckt.
  heatmaps <- plot_list$heatmap
  
  if (length(heatmaps)) {
    for (year_plot in heatmaps) {
      if (inherits(year_plot, "ggplot")) {
        print(year_plot)
      }
    }
  }
}


# ---------- Niederschlagsauswertung ----------

if (do_precip) {
  
  precip_plot_data <- make_plot_data(
    raw_precip,
    "precip",
    selected_station
  )
  
  precip_plots <- plot_time_series(
    dat=precip_plot_data,
    parameter="precip",
    selected_station=selected_station
  )
  
  precip_plots$anomaly <- plot_monthly_anomaly(
    dat=precip_plot_data,
    parameter="precip",
    selected_station=selected_station
  )
  
  precip_plots$heatmap <- plot_weather_heatmap(
    precip_plot_data,
    "precip",
    selected_station
  )
  
  save_weather_pdf(
    file.path(
      out_dir,
      paste0(
        "Niederschlag_",
        selected_station$Stations_id,
        "_",
        date_tag,
        ".pdf"
      )
    ),
    precip_plots
  )
}


# ---------- Temperaturauswertung ----------

if (do_temp) {
  
  temp_plot_data <- make_plot_data(
    raw_temp,
    "temp",
    selected_station
  )
  
  temp_plots <- plot_time_series(
    temp_plot_data,
    "temp",
    selected_station
  )
  
  temp_plots$anomaly <- plot_monthly_anomaly(
    temp_plot_data,
    "temp",
    selected_station
  )
  
  temp_plots$heatmap <- plot_weather_heatmap(
    temp_plot_data,
    "temp",
    selected_station
  )
  
  save_weather_pdf(
    file.path(
      out_dir,
      paste0(
        "Temperatur_",
        selected_station$Stations_id,
        "_",
        date_tag,
        ".pdf"
      )
    ),
    temp_plots
  )
}


# ---------- Abschlussmeldung ----------

message(
  "Fertig. PDF-Dateien liegen im Ordner: ",
  normalizePath(
    out_dir,
    winslash = "/",
    mustWork = FALSE
  )
)
