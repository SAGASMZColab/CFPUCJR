# Instala as dependências do Sistema Contábil PUC Jr
options(repos = c(CRAN = "https://cloud.r-project.org"))

pkgs <- c(
  "shiny",
  "bslib",
  "DT",
  "tidyverse",
  "googledrive",
  "googlesheets4",
  "gargle",
  "writexl",
  "openssl"
)

instalar <- pkgs[!pkgs %in% rownames(installed.packages())]
if (length(instalar) > 0) install.packages(instalar)

cat("Dependências instaladas com sucesso.\n")

