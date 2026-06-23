# =============================================================================
#  SISTEMA CONTÁBIL - PUC JUNIOR CONSULTORIA (Empresa Júnior)
#  Backend: Google Sheets + Google Drive  |  Login: Conta Google (Gmail)
#  Stack: R + Shiny + bslib
#
#  Desenvolvedores/Proprietários: guilhezago@gmail.com / staszago@gmail.com
#
#  ARQUIVO ÚNICO (app.R) - pronto para deploy em container/Shiny Server.
#  Leia o cabeçalho de CONFIGURAÇÃO abaixo e o guia "setup_google_vercel.md".
# =============================================================================

library(shiny)
library(bslib)
library(DT)
library(tidyverse)
library(googledrive)
library(googlesheets4)
library(gargle)        # service account p/ backend
library(writexl)       # exportação/backup local em .xlsx
# Hash de senha (openssl já vem como dependência do httr/gargle)

# =============================================================================
#  >>> CONFIGURAÇÃO (preencha com os seus dados) <<<
# =============================================================================
# Tudo abaixo pode ser definido por variáveis de ambiente (recomendado em
# produção) OU diretamente aqui durante o desenvolvimento.

CFG <- list(
  # --- MODO DE EXECUÇÃO ---------------------------------------------------
  # TRUE  = roda 100% local/em memória, com login simplificado (para
  #         desenvolver e testar SEM depender do Google).
  # FALSE = produção: login pela conta Google + backend no Google Sheets/Drive.
  MODO_DEMO = as.logical(Sys.getenv("PUCJR_MODO_DEMO", "TRUE")),

  # --- IDENTIDADE DOS ADMINISTRADORES (papel "contador") ------------------
  # Estes e-mails sempre entram como contador, mesmo que ainda não estejam
  # cadastrados na planilha de usuários (bootstrap do sistema).
  ADMINS = c("guilhezago@gmail.com", "staszago@gmail.com"),

  # --- GOOGLE SHEETS (backend de dados) -----------------------------------
  # ID da planilha-base (a parte entre /d/ e /edit na URL do Google Sheets).
  SHEET_ID = Sys.getenv("PUCJR_SHEET_ID", ""),

  # --- GOOGLE DRIVE (documentos de suporte) -------------------------------
  # ID da pasta do Drive onde os comprovantes serão salvos.
  DRIVE_FOLDER_ID = Sys.getenv("PUCJR_DRIVE_FOLDER_ID", ""),

  # --- SERVICE ACCOUNT (chave JSON p/ ler/gravar Sheets e Drive) ----------
  SA_JSON = Sys.getenv("PUCJR_SA_JSON", "service-account.json"),

  # Alternativa: conteúdo da chave JSON colado numa variável de ambiente
  # (útil em plataformas que só aceitam segredos como variável, ex.: Posit
  # Connect Cloud). Se preenchida, tem prioridade sobre SA_JSON.
  SA_JSON_CONTENT = Sys.getenv("PUCJR_SA_JSON_CONTENT", ""),

  # --- C+: TOKEN OAuth de uma conta COM cota no Drive (pessoal/dedicada) ----
  # Token gerado uma vez localmente (login no navegador) e codificado em
  # base64. É a única forma de o app CRIAR arquivos no Drive usando Gmail
  # comum (a conta de serviço não tem cota). Tem PRIORIDADE sobre a chave.
  GOOGLE_TOKEN_B64 = Sys.getenv("PUCJR_GOOGLE_TOKEN_B64", "")
)

# Nomes das abas (worksheets) usadas no Google Sheets
ABAS <- list(
  usuarios       = "usuarios",
  plano          = "plano_contas",
  diario         = "diario",
  saldos         = "saldos_iniciais",
  encerramentos  = "encerramentos"
)

# Contas usadas internamente pelo sistema (carga inicial e encerramento).
# Não podem ser excluídas do plano de contas.
CONTAS_SISTEMA <- c(
  "1.1.1.01.XX",  # Caixa (contrapartida de ajuste na carga inicial)
  "2.3.1.XX.XX",  # Patrimônio Social (contrapartida da carga inicial)
  "2.3.3.XX.XX"   # Superávit/Déficit do Período (encerramento)
)

# (O login é feito por e-mail + senha, gerenciado pelo próprio sistema —
#  ver schema_usuarios e a seção de LOGIN no servidor.)

# =============================================================================
#  FUNÇÕES AUXILIARES
# =============================================================================

# Operador "ou nulo" (compatibilidade entre versões)
`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || (length(a) == 1 && is.na(a))) b else a

# Formatação de moeda BRL
formatar_moeda <- function(x) {
  if (is.null(x) || length(x) == 0 || all(is.na(x))) return("R$ 0,00")
  x[is.na(x)] <- 0
  paste0("R$ ", format(round(x, 2), nsmall = 2, decimal.mark = ",", big.mark = ".", scientific = FALSE))
}

# Converte string "R$ 1.234,56" de volta para número
desformatar_moeda <- function(s) {
  s |>
    str_replace_all("R\\$\\s*", "") |>
    str_replace_all("\\.", "") |>
    str_replace_all(",", ".") |>
    as.numeric()
}

# Exercício (ano fiscal) a partir de uma data
exercicio_de <- function(d) as.integer(format(as.Date(d), "%Y"))

# Hash de senha (SHA-256) — nunca guardamos a senha em texto puro
hash_senha <- function(s) {
  # Normaliza para evitar divergência entre ambientes (Windows x Linux) e
  # espaços acidentais: força UTF-8, apara espaços das pontas e usa hex minúsculo.
  s <- trimws(enc2utf8(as.character(s %||% "")))
  if (!nzchar(s)) return("")
  tolower(as.character(openssl::sha256(charToRaw(s))))
}

# --- Validação de CPF / CNPJ ------------------------------------------------
validar_cpf <- function(x) {
  x <- gsub("\\D", "", x)
  if (nchar(x) != 11) return(FALSE)
  d <- as.integer(strsplit(x, "")[[1]])
  if (length(unique(d)) == 1) return(FALSE)            # rejeita 000..., 111...
  d1 <- (sum(d[1:9] * 10:2) * 10) %% 11; if (d1 == 10) d1 <- 0
  if (d1 != d[10]) return(FALSE)
  d2 <- (sum(d[1:10] * 11:2) * 10) %% 11; if (d2 == 10) d2 <- 0
  d2 == d[11]
}

validar_cnpj <- function(x) {
  x <- gsub("\\D", "", x)
  if (nchar(x) != 14) return(FALSE)
  d <- as.integer(strsplit(x, "")[[1]])
  if (length(unique(d)) == 1) return(FALSE)
  r1 <- sum(d[1:12] * c(5,4,3,2,9,8,7,6,5,4,3,2)) %% 11
  d1 <- if (r1 < 2) 0 else 11 - r1
  if (d1 != d[13]) return(FALSE)
  r2 <- sum(d[1:13] * c(6,5,4,3,2,9,8,7,6,5,4,3,2)) %% 11
  d2 <- if (r2 < 2) 0 else 11 - r2
  d2 == d[14]
}

formatar_cpf  <- function(x) sub("(\\d{3})(\\d{3})(\\d{3})(\\d{2})",
                                 "\\1.\\2.\\3-\\4", gsub("\\D", "", x))
formatar_cnpj <- function(x) sub("(\\d{2})(\\d{3})(\\d{3})(\\d{4})(\\d{2})",
                                 "\\1.\\2.\\3/\\4-\\5", gsub("\\D", "", x))

# Analisa um documento e devolve status/tipo/formatado
analisar_doc <- function(x) {
  d <- gsub("\\D", "", x %||% "")
  if (nchar(d) == 0) return(list(status = "vazio", tipo = NA, fmt = ""))
  if (nchar(d) == 11)
    return(list(status = if (validar_cpf(d)) "valido" else "invalido",
                tipo = "CPF", fmt = formatar_cpf(d)))
  if (nchar(d) == 14)
    return(list(status = if (validar_cnpj(d)) "valido" else "invalido",
                tipo = "CNPJ", fmt = formatar_cnpj(d)))
  list(status = "invalido", tipo = NA, fmt = x)
}

# =============================================================================
#  PLANO DE CONTAS PADRÃO (semente / fallback)
#  Gerado a partir de "Plano de Contas.xlsx" (aba PC) - 63 contas.
#  Em produção, o plano vivo fica no Google Sheets e é editável pelo contador.
# =============================================================================

db_plano_contas <- tibble(
  Codigo = c(
    "1.1.1.01.XX",
    "1.1.1.02.01",
    "1.1.1.02.02",
    "1.1.1.03.XX",
    "1.1.2.01.01",
    "1.1.2.01.02",
    "1.1.2.02.XX",
    "1.1.2.03.XX",
    "1.1.2.04.XX",
    "1.1.2.05.XX",
    "1.1.3.01.XX",
    "1.1.3.02.XX",
    "1.1.3.03.XX",
    "1.2.1.01.XX",
    "1.2.2.01.XX",
    "1.2.2.02.XX",
    "1.2.3.01.XX",
    "1.2.3.02.XX",
    "1.2.3.03.XX",
    "1.2.3.04.XX",
    "1.2.3.05.XX",
    "1.2.3.06.XX",
    "1.2.4.XX.XX",
    "1.2.5.01.XX",
    "1.2.5.02.XX",
    "2.1.1.01.XX",
    "2.1.1.02.XX",
    "2.1.1.02.01",
    "2.1.1.02.02",
    "2.1.1.03.01",
    "2.1.1.03.02",
    "2.1.1.04.XX",
    "2.1.2.01.XX",
    "2.1.2.02.XX",
    "2.1.3.01.XX",
    "2.2.1.XX.XX",
    "2.2.2.XX.XX",
    "2.3.1.XX.XX",
    "2.3.2.XX.XX",
    "2.3.3.XX.XX",
    "3.1.1.01.XX",
    "3.1.1.02.XX",
    "3.1.2.01.XX",
    "3.1.2.02.XX",
    "3.1.2.03.XX",
    "3.1.2.04.XX",
    "3.1.3.XX.XX",
    "3.1.4.XX.XX",
    "3.2.1.XX.XX",
    "3.3.1.XX.XX",
    "3.4.X.XX.XX",
    "3.5.1.XX.XX",
    "4.1.1.01.XX",
    "4.1.1.02.XX",
    "4.1.2.XX.XX",
    "4.1.3.XX.XX",
    "4.2.1.XX.XX",
    "4.2.2.XX.XX",
    "4.3.1.XX.XX",
    "4.4.1.XX.XX",
    "5.1.1.01.XX",
    "5.1.2.XX.XX",
    "5.2.1.XX.XX"
  ),
  Nome = c(
    "Caixa",
    "Banco - Depósitos Bancários à Vista - Recursos Livres",
    "Banco - Depósitos Bancários à Vista - Recursos com Restrição",
    "Aplicações Financeiras de Liquidez Imediata",
    "Contas a Receber - Créditos de Mensalidades",
    "Contas a Receber - Créditos de Serviços Prestados",
    "(-)  Perda Estimada para Créditos de Liquidação Duvidosa",
    "Adiantamentos a Empregados",
    "Adiantamento a Fornecedores",
    "Despesas Antecipadas",
    "Estoque para Revenda",
    "Almoxarifado",
    "(-) Ajuste a Valor Recuperável (AVR)",
    "Créditos de  Longo Prazo",
    "Investimentos - Participações Societárias",
    "Propriedades para Investimento",
    "Imóveis de Uso",
    "Móveis e Utensílios",
    "Equipamentos",
    "Veículos",
    "Instalações",
    "Obras em Andamento",
    "(-) Depreciação Acumulada",
    "Softwares",
    "(-) Amortização Acumulada",
    "Contas a Pagar - Fornecedores",
    "Contas a Pagar - Obrigações com Empregados",
    "Salários a Pagar",
    "Encargos Sociais a Pagar",
    "Obrigações Tributárias - Impostos a Recolher",
    "Obrigações Tributárias - Contribuições Sociais a Recolher",
    "Empréstimos e Financiamentos (Curto Prazo)",
    "Convênios e Parcerias em Execução",
    "Termos de Parceria em Execução",
    "Provisões para Demandas Judiciais",
    "Empréstimos e Financiamentos (Longo Prazo)",
    "Parcerias com Restrição – Longo Prazo",
    "Patrimônio Social",
    "Reservas",
    "Superávit ou Déficit do Período",
    "Salários",
    "Encargos Sociais",
    "Aluguel",
    "Energia Elétrica",
    "Água e Esgoto",
    "Material de Consumo",
    "Despesas Tributárias",
    "Depreciação e Amortização",
    "Despesas de Projetos",
    "Gratuidades Concedidas",
    "Serviços Voluntários",
    "Outras Despesas",
    "Doações Voluntárias",
    "Subvenções",
    "Receita de Serviços Prestados",
    "Receitas de Projetos sem Restrição",
    "Receitas de Projetos com Restrição -Convênios Públicos",
    "Receitas de Projetos com Restrição - Parcerias Privadas",
    "Rendimentos de Aplicações Financeiras",
    "Outras Receitas",
    "Benefícios Obtidos - Isenção de Tributos",
    "Benefícios Obtidos - Serviços Voluntários Obtidos",
    "Benefícios Concedidos - Gratuidade Concedida"
  ),
  Subgrupo = c(
    "Caixa e Equivalentes de Caixa",
    "Caixa e Equivalentes de Caixa",
    "Caixa e Equivalentes de Caixa",
    "Caixa e Equivalentes de Caixa",
    "Créditos",
    "Créditos",
    "Créditos",
    "Créditos",
    "Créditos",
    "Créditos",
    "Estoques",
    "Estoques",
    "Estoques",
    "Realizável a Longo Prazo",
    "Investimentos",
    "Investimentos",
    "Imobilizado",
    "Imobilizado",
    "Imobilizado",
    "Imobilizado",
    "Imobilizado",
    "Imobilizado",
    "Imobilizado",
    "Intangível",
    "Intangível",
    "Contas a Pagar",
    "Contas a Pagar",
    "Contas a Pagar",
    "Contas a Pagar",
    "Contas a Pagar",
    "Contas a Pagar",
    "Contas a Pagar",
    "Parcerias com Restrição",
    "Parcerias com Restrição",
    "Provisões",
    "Passivo Não Circulante",
    "Passivo Não Circulante",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Pessoal Administrativo",
    "Pessoal Administrativo",
    "Serviços Gerais",
    "Serviços Gerais",
    "Serviços Gerais",
    "Serviços Gerais",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas de Projetos",
    "Gratuidades Concedidas",
    "Serviços Voluntários",
    "Outras Despesas",
    "Doações e Contribuições",
    "Doações e Contribuições",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas Financeiras",
    "Outras Receitas",
    "Renúncia Fiscal Obtida",
    "Serviços Voluntários Obtidos",
    "Benefícios Concedidos - Gratuidade Concedida"
  ),
  Grupo_Nivel2 = c(
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Ativo Não Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Circulante",
    "Passivo Não Circulante",
    "Passivo Não Circulante",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas Administrativas",
    "Despesas de Projetos",
    "Gratuidades Concedidas",
    "Serviços Voluntários",
    "Outras Despesas",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas de Operações Próprias",
    "Receitas de Projetos com Restrição",
    "Receitas de Projetos com Restrição",
    "Receitas Financeiras",
    "Outras Receitas",
    "Benefícios Obtidos",
    "Benefícios Obtidos",
    "Benefícios Concedidos - Gratuidade"
  ),
  Grupo = c(
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Ativo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Passivo",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Patrimônio Líquido",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Despesas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Receitas",
    "Variações Patrimoniais",
    "Variações Patrimoniais",
    "Variações Patrimoniais"
  ),
  Tipo = c(
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Patrimonial",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado",
    "Resultado"
  )
)

# Coluna auxiliar de nível (qtde de segmentos) e ordenação
db_plano_contas <- db_plano_contas %>%
  mutate(Nivel = str_count(Codigo, "\\.") + 1) %>%
  arrange(Codigo)

# =============================================================================
#  ESQUEMAS (estruturas vazias) DAS TABELAS DE DADOS
# =============================================================================
schema_diario <- tibble(
  ID = numeric(), Lancamento_ID = numeric(), Data = as.Date(character()),
  Conta_Debito = character(), Conta_Credito = character(),
  Valor = numeric(), Historico = character(), Doc_Link = character(),
  Tipo_Lancamento = character(), Ref_ID = numeric(),
  Exercicio = integer(), Usuario = character(), Timestamp = character()
)

schema_saldos <- tibble(
  Codigo = character(), Nome = character(),
  Saldo_Inicial = numeric(), Data_Base = as.Date(character())
)

schema_usuarios <- tibble(
  email = character(), nome = character(), papel = character(),
  ativo = logical(), senha_hash = character(), criado_em = character()
)

schema_encerramentos <- tibble(
  Exercicio = integer(), Data_Encerramento = as.Date(character()),
  Total_Receitas = numeric(), Total_Despesas = numeric(),
  Resultado = numeric(), Arquivo_Link = character(),
  Usuario = character(), Timestamp = character()
)

# Usuários-semente (os dois proprietários como contador). A senha fica em
# branco: cada um define a própria no primeiro acesso (sem segredo no código).
seed_usuarios <- tibble(
  email = CFG$ADMINS,
  nome  = c("Guilherme (Dev)", "Stas (Dev)")[seq_along(CFG$ADMINS)],
  papel = "contador",
  ativo = TRUE,
  senha_hash = "",
  criado_em = as.character(Sys.time())
)

# =============================================================================
#  CAMADA DE BACKEND (Google Sheets / Drive)  -- abstração com fallback
# =============================================================================

# Autentica o backend (apenas em produção)
backend_conectar <- function() {
  if (CFG$MODO_DEMO) return(invisible(FALSE))
  scopes <- c("https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive")
  ok <- tryCatch({
    # Monta o token base64: usa PUCJR_GOOGLE_TOKEN_B64 inteiro; se vazio,
    # concatena as partes PUCJR_GOOGLE_TOKEN_B64_1, _2, ... (para contornar o
    # limite de tamanho por variável do Connect Cloud).
    tok_b64 <- CFG$GOOGLE_TOKEN_B64
    if (!nzchar(tok_b64)) {
      partes <- character(0); i <- 1L
      repeat {
        v <- Sys.getenv(paste0("PUCJR_GOOGLE_TOKEN_B64_", i), "")
        if (!nzchar(v)) break
        partes <- c(partes, v); i <- i + 1L
      }
      if (length(partes)) tok_b64 <- paste(partes, collapse = "")
    }
    tok_b64 <- gsub("[[:space:]]", "", tok_b64)  # remove quebras/espaços acidentais

    if (nzchar(tok_b64)) {
      # C+: token OAuth de uma conta COM cota (lido de variável secreta).
      # Permite CRIAR arquivos no Drive (upload de comprovantes) com Gmail comum.
      raw <- openssl::base64_decode(tok_b64)
      tf  <- file.path(tempdir(), "pucjr-oauth-token.rds")
      writeBin(raw, tf)
      tok <- readRDS(tf)
      # O token foi gerado localmente com cache em ".secrets" — caminho que NÃO existe
      # no servidor. Ao renovar o token, o gargle tenta ler/gravar esse cache e falha.
      # Solução: redirecionar o cache para um diretório TEMPORÁRIO gravável e gravar o
      # token lá, para qualquer leitura/escrita do gargle funcionar.
      cache_dir <- file.path(tempdir(), "gargle_cache")
      dir.create(cache_dir, showWarnings = FALSE, recursive = TRUE)
      fname <- tryCatch({
        bp <- basename(tok$cache_path)
        if (is.character(bp) && length(bp) == 1 && nzchar(bp)) bp else "pucjr_token"
      }, error = function(e) "pucjr_token")
      novo_cache <- file.path(cache_dir, fname)
      try(tok$cache_path <- novo_cache, silent = TRUE)
      try(saveRDS(tok, novo_cache), silent = TRUE)
      options(gargle_oauth_cache = cache_dir)
      drive_auth(token = tok)
      gs4_auth(token = drive_token())
      message("Autenticado via token OAuth [cache-tmp v2] (conta com cota no Drive). Partes: ",
              if (nzchar(CFG$GOOGLE_TOKEN_B64)) "1 (variável única)" else (i - 1L))
    } else {
      # Se o conteúdo da chave veio por variável de ambiente, grava num arquivo
      # temporário e usa esse caminho (caso típico do Posit Connect Cloud).
      sa_path <- CFG$SA_JSON
      if (nzchar(CFG$SA_JSON_CONTENT)) {
        sa_path <- file.path(tempdir(), "pucjr-sa.json")
        writeLines(CFG$SA_JSON_CONTENT, sa_path)
      }
      if (nzchar(sa_path) && file.exists(sa_path)) {
        # Conta de serviço por chave JSON (grava em Sheets; NÃO cria arquivos
        # no Drive com Gmail comum por falta de cota).
        gs4_auth(path = sa_path)
        drive_auth(path = sa_path)
      } else {
        # Sem chave: Application Default Credentials / metadados / Federação.
        token <- gargle::token_fetch(scopes = scopes)
        if (is.null(token))
          stop("Nenhuma credencial encontrada (configure o token OAuth, a conta de serviço anexada ou GOOGLE_APPLICATION_CREDENTIALS).")
        gs4_auth(token = token)
        drive_auth(token = token)
      }
    }
    TRUE
  }, error = function(e) {
    message("Falha ao autenticar o backend: ", e$message); FALSE
  })
  invisible(ok)
}

# Lê uma aba; em caso de erro/DEMO devolve o schema vazio recebido
backend_ler <- function(aba, schema) {
  if (CFG$MODO_DEMO || !nzchar(CFG$SHEET_ID)) return(schema)
  tryCatch({
    df <- read_sheet(CFG$SHEET_ID, sheet = aba, col_types = "c")
    if (nrow(df) == 0) return(schema)
    # Reconverte tipos conforme o schema
    for (col in names(schema)) {
      if (!col %in% names(df)) df[[col]] <- NA
      cls <- class(schema[[col]])[1]
      df[[col]] <- switch(cls,
        numeric = suppressWarnings(as.numeric(df[[col]])),
        integer = suppressWarnings(as.integer(df[[col]])),
        logical = as.logical(df[[col]]),
        Date    = suppressWarnings(as.Date(df[[col]])),
        as.character(df[[col]])
      )
    }
    df[, names(schema)]
  }, error = function(e) {
    message("backend_ler(", aba, "): ", e$message); schema
  })
}

# Sobrescreve uma aba inteira (idempotente; ideal p/ esta escala de dados)
backend_gravar <- function(aba, df) {
  if (CFG$MODO_DEMO || !nzchar(CFG$SHEET_ID)) return(invisible(TRUE))
  tryCatch({
    out <- df %>% mutate(across(where(is.Date), as.character))
    sheet_write(out, ss = CFG$SHEET_ID, sheet = aba)
    TRUE
  }, error = function(e) {
    message("backend_gravar(", aba, "): ", e$message); FALSE
  })
}

# Faz upload de um comprovante ao Drive e devolve o link compartilhável
backend_upload_doc <- function(caminho, nome) {
  if (CFG$MODO_DEMO || !nzchar(CFG$DRIVE_FOLDER_ID)) {
    return(paste0("[DEMO] ", nome, " (upload simulado)"))
  }
  tryCatch({
    dest <- if (nzchar(CFG$DRIVE_FOLDER_ID)) as_id(CFG$DRIVE_FOLDER_ID) else NULL
    # Nome único no Drive (evita colisão "já existe": data/hora + sufixo aleatório)
    nome_unico <- paste0(format(Sys.time(), "%Y%m%d-%H%M%S"), "-",
                         formatC(sample.int(9999, 1), width = 4, flag = "0"),
                         "_", nome)
    arq <- drive_upload(media = caminho, path = dest, name = nome_unico, overwrite = FALSE)
    drive_share(arq, role = "reader", type = "anyone")
    arq$drive_resource[[1]]$webViewLink %||% drive_link(arq)
  }, error = function(e) {
    message("upload Drive: ", e$message)
    paste0("Falha no upload: ", nome)
  })
}

# Garante que a planilha-base tenha as abas/cabeçalhos (executar 1x no deploy)
backend_inicializar <- function() {
  if (CFG$MODO_DEMO || !nzchar(CFG$SHEET_ID)) return(invisible(FALSE))
  abas_existentes <- tryCatch(sheet_names(CFG$SHEET_ID), error = function(e) character())
  semear <- list(
    usuarios      = seed_usuarios,
    plano_contas  = db_plano_contas,
    diario        = schema_diario,
    saldos_iniciais = schema_saldos,
    encerramentos = schema_encerramentos
  )
  for (aba in names(semear)) {
    if (!aba %in% abas_existentes) {
      backend_gravar(aba, semear[[aba]])
    }
  }
  invisible(TRUE)
}

# =============================================================================
#  LÓGICA CONTÁBIL COMPARTILHADA
# =============================================================================

# Calcula o saldo de cada conta a partir do diário, respeitando a natureza
saldos_por_conta <- function(diario, plano) {
  if (nrow(diario) == 0) {
    return(plano %>% mutate(Saldo = 0))
  }
  deb <- diario %>% group_by(Codigo = Conta_Debito) %>%
    summarise(D = sum(Valor, na.rm = TRUE), .groups = "drop")
  cre <- diario %>% group_by(Codigo = Conta_Credito) %>%
    summarise(C = sum(Valor, na.rm = TRUE), .groups = "drop")
  plano %>%
    left_join(deb, by = "Codigo") %>%
    left_join(cre, by = "Codigo") %>%
    mutate(
      D = replace_na(D, 0), C = replace_na(C, 0),
      Saldo = case_when(
        str_starts(Codigo, "1")   ~ D - C,   # Ativo (devedora)
        str_starts(Codigo, "2")   ~ C - D,   # Passivo + PL (credora)
        str_starts(Codigo, "3")   ~ D - C,   # Despesa (devedora)
        str_starts(Codigo, "4")   ~ C - D,   # Receita (credora)
        str_starts(Codigo, "5.2") ~ D - C,   # Benefícios concedidos
        str_starts(Codigo, "5")   ~ C - D,   # Benefícios obtidos
        TRUE ~ D - C
      )
    ) %>%
    select(-D, -C)
}

# Dados tabulares do Balanço Patrimonial (posição acumulada até 'ate_ano').
# Se ate_ano = NULL, usa todo o diário (posição atual).
dados_balanco <- function(diario, plano, ate_ano = NULL) {
  d <- diario
  if (!is.null(ate_ano)) d <- d %>% filter(Exercicio <= ate_ano)
  saldos <- saldos_por_conta(d, plano)

  # Patrimoniais, EXCETO 2.3.3 (Superávit/Déficit do Período) — tratada à parte
  s <- saldos %>%
    filter(Tipo == "Patrimonial", !str_starts(Codigo, "2.3.3"), abs(Saldo) > 0.01) %>%
    mutate(Grupo_BP = if_else(Grupo == "Ativo", "ATIVO",
                              "PASSIVO + PATRIMÔNIO LÍQUIDO")) %>%
    arrange(Codigo) %>%
    select(Grupo = Grupo_BP, Grupo_Nivel2, Codigo, Conta = Nome, Saldo)

  # Conta 2.3.3 = resultado(s) de exercício(s) já encerrado(s) acumulado(s) na
  # própria conta (saldo real) + resultado do exercício corrente ainda em aberto
  # (saldos das contas de resultado 3/4/5). O resultado da DRE PERMANECE em
  # 2.3.3 — não é transferido para o Patrimônio Social (2.3.1). Assim o Balanço
  # fecha (Ativo = Passivo + PL) tanto durante o exercício quanto após encerrar.
  saldo_233 <- saldos %>% filter(str_starts(Codigo, "2.3.3")) %>%
    summarise(v = sum(Saldo, na.rm = TRUE)) %>% pull(v)
  if (length(saldo_233) == 0 || is.na(saldo_233)) saldo_233 <- 0
  res <- saldos %>% filter(Tipo == "Resultado")
  resultado_corrente <- if (nrow(res) == 0) 0 else sum(
    ifelse(str_starts(res$Codigo, "3") | str_starts(res$Codigo, "5.2"),
           -res$Saldo, res$Saldo), na.rm = TRUE)
  total_233 <- saldo_233 + resultado_corrente
  if (abs(total_233) > 0.01) {
    cr <- plano %>% filter(str_starts(Codigo, "2.3.3"))
    nome_res <- if (nrow(cr) > 0) cr$Nome[1] else "Superávit ou Déficit do Período"
    cod_res  <- if (nrow(cr) > 0) cr$Codigo[1] else "2.3.3.XX.XX"
    s <- bind_rows(s, tibble(
      Grupo = "PASSIVO + PATRIMÔNIO LÍQUIDO",
      Grupo_Nivel2 = "Patrimônio Líquido",
      Codigo = cod_res, Conta = nome_res, Saldo = total_233))
  }
  s
}

# Dados tabulares da DRE de um exercício (exclui lançamentos de encerramento)
dados_dre <- function(diario, plano, ano) {
  d <- diario %>% filter(Exercicio == ano,
            is.na(Tipo_Lancamento) | Tipo_Lancamento != "Encerramento")
  s <- saldos_por_conta(d, plano)
  rec  <- s %>% filter(str_starts(Codigo, "4"), abs(Saldo) > 0.01) %>%
    mutate(Categoria = "Receita")
  desp <- s %>% filter(str_starts(Codigo, "3"), abs(Saldo) > 0.01) %>%
    mutate(Categoria = "Despesa")
  bind_rows(rec, desp) %>% arrange(Codigo) %>%
    select(Categoria, Codigo, Conta = Nome, Saldo)
}

# Relatório de pagamentos/recebimentos por CNPJ/CPF, a partir das
# movimentações de Caixa e Equivalentes (contas iniciadas em "1.1.1").
#   lado = "pagamento"  -> crédito em caixa/banco (saídas)
#   lado = "recebimento"-> débito  em caixa/banco (entradas)
# O CNPJ/CPF e o nome são extraídos do campo Histórico.
relatorio_caixa_terceiros <- function(diario, plano, lado, ano = NULL) {
  cash <- plano$Codigo[str_starts(plano$Codigo, "1.1.1")]
  d <- diario %>% filter(Tipo_Lancamento %in% c("Movimento", "Estorno"))
  if (!is.null(ano) && ano != "Todos")
    d <- d %>% filter(Exercicio == suppressWarnings(as.integer(ano)))
  d <- if (lado == "pagamento") d %>% filter(Conta_Credito %in% cash)
       else                     d %>% filter(Conta_Debito %in% cash)
  if (nrow(d) == 0) return(tibble())
  d %>%
    mutate(
      Documento = coalesce(
        str_extract(Historico, "\\d{2}\\.\\d{3}\\.\\d{3}/\\d{4}-\\d{2}"),  # CNPJ
        str_extract(Historico, "\\d{3}\\.\\d{3}\\.\\d{3}-\\d{2}")),        # CPF
      Nome = str_match(Historico, "^\\s*(.*?)\\s*—\\s*(?:CNPJ|CPF):")[, 2]
    ) %>%
    mutate(
      Documento = replace_na(Documento, "(sem identificação)"),
      Nome = ifelse(is.na(Nome) | trimws(Nome) == "", "—", trimws(Nome))
    ) %>%
    group_by(Documento, Nome) %>%
    summarise(`Nº lançamentos` = n(), Total = sum(Valor, na.rm = TRUE),
              .groups = "drop") %>%
    arrange(desc(Total))
}

# Cria a planilha arquivada do exercício no Google Drive e devolve o link.
# Em DEMO, devolve mensagem simulada (sem criar arquivo).
backend_arquivar_exercicio <- function(ano, abas) {
  if (CFG$MODO_DEMO || !nzchar(CFG$SHEET_ID)) {
    return(paste0("[DEMO] PUCJr_Encerramento_", ano, " (arquivo simulado)"))
  }
  tryCatch({
    nome <- paste0("PUCJr_Encerramento_", ano)
    ss <- gs4_create(nome, sheets = abas)
    if (nzchar(CFG$DRIVE_FOLDER_ID)) {
      drive_mv(googledrive::as_id(ss), path = googledrive::as_id(CFG$DRIVE_FOLDER_ID))
    }
    drive_share(googledrive::as_id(ss), role = "reader", type = "anyone")
    paste0("https://docs.google.com/spreadsheets/d/", as.character(ss), "/edit")
  }, error = function(e) {
    message("arquivar_exercicio(", ano, "): ", e$message)
    paste0("Falha ao arquivar exercício ", ano)
  })
}

# =============================================================================
#  MÓDULO: CARGA INICIAL DE SALDOS
# =============================================================================
cargaSaldosUI <- function(id) {
  ns <- NS(id)
  layout_sidebar(
    sidebar = sidebar(
      title = "Configuração da Carga", width = 330,
      dateInput(ns("data_base"), "Data Base do Balanço:", value = Sys.Date() - 30),
      helpText("Saldos de abertura das contas patrimoniais (Ativo e Passivo + PL)."),
      selectInput(ns("tipo_conta"), "Filtrar:",
                  choices = c("Todas", "Ativo (1)", "Passivo + PL (2)")),
      hr(),
      h6("Importação em lote"),
      fileInput(ns("arquivo_csv"), "CSV de saldos", accept = ".csv",
                buttonLabel = "Selecionar", placeholder = "Codigo;Saldo_Inicial"),
      hr(),
      actionButton(ns("btn_aplicar"), "Aplicar saldos no sistema",
                   class = "btn-success w-100", icon = icon("check-circle")),
      br(), br(),
      actionButton(ns("btn_limpar"), "Limpar saldos iniciais",
                   class = "btn-outline-danger w-100", icon = icon("trash-alt"))
    ),
    card(
      card_header("Saldos iniciais por conta patrimonial",
                  tags$span(class = "badge bg-info ms-2", "edite a coluna Saldo_Inicial")),
      DTOutput(ns("tabela"))
    ),
    card(card_header("Resumo da carga"), uiOutput(ns("resumo")))
  )
}

cargaSaldosServer <- function(id, plano, diario, saldos, persistir) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    temp <- reactiveVal(NULL)

    contas_patrim <- reactive({
      req(plano())
      plano() %>% filter(Tipo == "Patrimonial")
    })

    observe({
      req(contas_patrim())
      df <- contas_patrim()
      if (!is.null(saldos()) && nrow(saldos()) > 0) {
        df <- df %>%
          left_join(saldos() %>% select(Codigo, Saldo_Inicial), by = "Codigo") %>%
          mutate(Saldo_Inicial = replace_na(Saldo_Inicial, 0))
      } else {
        df <- df %>% mutate(Saldo_Inicial = 0)
      }
      if (!is.null(input$tipo_conta)) {
        if (input$tipo_conta == "Ativo (1)")        df <- df %>% filter(str_starts(Codigo, "1"))
        if (input$tipo_conta == "Passivo + PL (2)") df <- df %>% filter(str_starts(Codigo, "2"))
      }
      temp(df)
    })

    output$tabela <- renderDT({
      req(temp())
      df <- temp()
      if (nrow(df) == 0)
        return(datatable(tibble(Mensagem = "Sem contas"), options = list(dom = "t"), rownames = FALSE))
      df %>%
        mutate(Saldo_Formatado = formatar_moeda(Saldo_Inicial)) %>%
        select(Codigo, Nome, Saldo_Formatado, Saldo_Inicial) %>%
        datatable(rownames = FALSE,
          editable = list(target = "cell", disable = list(columns = c(0, 1, 2))),
          options = list(pageLength = 15, scrollX = TRUE)) %>%
        formatStyle("Saldo_Inicial",
          backgroundColor = styleInterval(0, c("#ffe6e6", "#e6ffe6")))
    }, server = FALSE)

    observeEvent(input$tabela_cell_edit, {
      info <- input$tabela_cell_edit
      df <- temp()
      if (!is.null(info) && nrow(df) > 0 && info$col == 3) {
        v <- suppressWarnings(as.numeric(info$value))
        if (!is.na(v) && info$row <= nrow(df)) {
          df[info$row, "Saldo_Inicial"] <- v
          temp(df)
        } else {
          showNotification("Valor inválido. Use números.", type = "warning")
        }
      }
    })

    observeEvent(input$arquivo_csv, {
      req(input$arquivo_csv)
      tryCatch({
        csv <- read.csv2(input$arquivo_csv$datapath, stringsAsFactors = FALSE)
        if (!all(c("Codigo", "Saldo_Inicial") %in% names(csv))) {
          showNotification("CSV precisa das colunas Codigo e Saldo_Inicial.", type = "error"); return()
        }
        df <- temp()
        for (i in seq_len(nrow(csv))) {
          cod <- as.character(csv$Codigo[i]); s <- suppressWarnings(as.numeric(csv$Saldo_Inicial[i]))
          if (!is.na(s) && cod %in% df$Codigo) df[df$Codigo == cod, "Saldo_Inicial"] <- s
        }
        temp(df)
        showNotification(paste("Importados", nrow(csv), "registros."), type = "message")
      }, error = function(e) showNotification(paste("Erro CSV:", e$message), type = "error"))
    })

    aplicar <- function(df_saldos) {
      ta <- df_saldos %>% filter(str_starts(Codigo, "1")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      tp <- df_saldos %>% filter(str_starts(Codigo, "2")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      diff <- tp - ta
      CONTA_PL <- "2.3.1.XX.XX"  # Patrimônio Social (contrapartida)

      novos <- tibble()
      for (i in seq_len(nrow(df_saldos))) {
        cod <- df_saldos$Codigo[i]; s <- df_saldos$Saldo_Inicial[i]
        if (abs(s) > 0.01) {
          if (str_starts(cod, "1")) {
            novos <- bind_rows(novos, tibble(
              Conta_Debito = cod, Conta_Credito = CONTA_PL, Valor = abs(s)))
          } else if (str_starts(cod, "2")) {
            if (s >= 0) novos <- bind_rows(novos, tibble(
              Conta_Debito = CONTA_PL, Conta_Credito = cod, Valor = abs(s)))
            else novos <- bind_rows(novos, tibble(
              Conta_Debito = cod, Conta_Credito = CONTA_PL, Valor = abs(s)))
          }
        }
      }
      # Ajuste automático se desbalanceado
      if (abs(diff) > 0.01) {
        if (diff > 0) novos <- bind_rows(novos, tibble(
          Conta_Debito = "1.1.1.01.XX", Conta_Credito = CONTA_PL, Valor = abs(diff)))
        else novos <- bind_rows(novos, tibble(
          Conta_Debito = CONTA_PL, Conta_Credito = "1.1.1.01.XX", Valor = abs(diff)))
      }

      d <- diario() %>% filter(is.na(Tipo_Lancamento) | Tipo_Lancamento != "Saldo_Inicial")
      if (nrow(novos) > 0) {
        prox <- if (nrow(d) == 0) 0 else max(d$ID, na.rm = TRUE)
        novos <- novos %>% mutate(
          ID = prox + row_number(),
          Data = input$data_base,
          Historico = paste("CARGA INICIAL - abertura em", format(input$data_base, "%d/%m/%Y")),
          Doc_Link = "Carga inicial do sistema",
          Tipo_Lancamento = "Saldo_Inicial", Ref_ID = NA_real_,
          Exercicio = exercicio_de(input$data_base),
          Usuario = "sistema", Timestamp = as.character(Sys.time())
        ) %>% select(all_of(names(schema_diario)))
        d <- bind_rows(d, novos)
      }
      diario(d); persistir("diario", d)

      sx <- df_saldos %>% select(Codigo, Nome, Saldo_Inicial) %>%
        mutate(Data_Base = input$data_base) %>% filter(abs(Saldo_Inicial) > 0.01)
      saldos(sx); persistir("saldos_iniciais", sx)

      showNotification(paste("Carga concluída:", nrow(novos), "lançamentos."),
                       type = "message", duration = 5)
    }

    observeEvent(input$btn_aplicar, {
      req(temp()); df <- temp()
      if (nrow(df) == 0) { showNotification("Nada a aplicar.", type = "warning"); return() }
      ta <- df %>% filter(str_starts(Codigo, "1")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      tp <- df %>% filter(str_starts(Codigo, "2")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      if (abs(ta - tp) > 0.01) {
        showModal(modalDialog(title = "Balanço desbalanceado",
          div(class = "alert alert-warning",
            p(paste("Ativo:", formatar_moeda(ta))),
            p(paste("Passivo + PL:", formatar_moeda(tp))),
            p(paste("Diferença:", formatar_moeda(abs(ta - tp)))),
            p("O sistema criará uma conta de ajuste. Continuar?")),
          footer = tagList(modalButton("Cancelar"),
            actionButton(ns("conf_aplicar"), "Continuar com ajuste", class = "btn-warning"))))
      } else aplicar(df)
    })
    observeEvent(input$conf_aplicar, { removeModal(); aplicar(temp()) })

    observeEvent(input$btn_limpar, {
      showModal(modalDialog(title = "Confirmar",
        "Remover TODOS os saldos iniciais e seus lançamentos de abertura?",
        footer = tagList(modalButton("Cancelar"),
          actionButton(ns("conf_limpar"), "Confirmar", class = "btn-danger"))))
    })
    observeEvent(input$conf_limpar, {
      removeModal()
      d <- diario() %>% filter(is.na(Tipo_Lancamento) | Tipo_Lancamento != "Saldo_Inicial")
      diario(d); persistir("diario", d)
      saldos(schema_saldos); persistir("saldos_iniciais", schema_saldos)
      df <- contas_patrim() %>% mutate(Saldo_Inicial = 0); temp(df)
      showNotification("Saldos iniciais removidos.", type = "message")
    })

    output$resumo <- renderUI({
      req(temp()); df <- temp()
      ta <- df %>% filter(str_starts(Codigo, "1")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      tp <- df %>% filter(str_starts(Codigo, "2")) %>% pull(Saldo_Inicial) %>% sum(na.rm = TRUE)
      ok <- abs(ta - tp) < 0.01
      div(class = "row",
        div(class = "col-md-4", value_box("Contas com saldo",
          sum(abs(df$Saldo_Inicial) > 0.01), showcase = icon("calculator"), theme = "info")),
        div(class = "col-md-4", value_box("Total do Ativo",
          formatar_moeda(ta), showcase = icon("chart-line"), theme = if (ok) "success" else "warning")),
        div(class = "col-md-4", value_box("Total Passivo + PL",
          formatar_moeda(tp), showcase = icon("scale-balanced"), theme = if (ok) "success" else "warning")),
        if (!ok) div(class = "col-12 mt-2", div(class = "alert alert-warning",
          icon("triangle-exclamation"), " Desbalanceado em ",
          formatar_moeda(abs(ta - tp)), ". Ao aplicar, será gerado ajuste automático."))
      )
    })
  })
}

# =============================================================================
#  INTERFACE (UI)
# =============================================================================
ui <- page_fluid(
  theme = bs_theme(
    version = 5, bootswatch = "flatly",
    primary = "#2c3e50", success = "#18bc9c",
    "navbar-padding-y" = "0.85rem",
    base_font = font_collection(
      "Inter", "Segoe UI", "system-ui", "-apple-system", "Helvetica Neue", "sans-serif"),
    heading_font = font_collection(
      "Inter", "Segoe UI", "system-ui", "-apple-system", "sans-serif")
  ),
  tags$head(tags$style(HTML("
    /* ---------- Cabeçalho / navbar ---------- */
    .navbar {
      padding-top: .85rem; padding-bottom: .85rem;
      box-shadow: 0 2px 10px rgba(0,0,0,.08);
      border-bottom: 3px solid #18bc9c;
    }
    .navbar-brand {
      font-weight: 700; font-size: 1.18rem;
      letter-spacing: .2px; padding-right: 1.5rem;
      display: flex; align-items: center; gap: .6rem;
    }
    .navbar .nav-link { padding: .55rem 1rem; font-weight: 500; }
    .navbar .nav-link.active { font-weight: 700; }
    .navbar-text { opacity: .85; font-size: .9rem; }

    /* ---------- Conteúdo ---------- */
    .bslib-page-navbar > .navbar + .container-fluid,
    .tab-content { padding-top: 1.4rem; }
    body { background: #f6f8fa; }

    /* ---------- Cartões ---------- */
    .card { border: none; border-radius: 12px;
            box-shadow: 0 1px 6px rgba(0,0,0,.06); margin-bottom: 1.1rem; }
    .card > .card-header {
      font-weight: 600; background: #fff; border-bottom: 1px solid #eef1f4;
      padding: .9rem 1.1rem;
    }
    .card-body { padding: 1.1rem; }

    /* ---------- Sidebar de formulário ---------- */
    .bslib-sidebar-layout > .sidebar { background: #fff; }
    .sidebar .form-label { font-weight: 600; font-size: .86rem; color: #2c3e50; }

    /* ---------- Subtítulo de seção no formulário ---------- */
    .form-section {
      font-size: .74rem; font-weight: 700; letter-spacing: .8px;
      text-transform: uppercase; color: #18bc9c;
      margin: 1rem 0 .4rem; padding-bottom: .25rem;
      border-bottom: 1px solid #e9edf1;
    }
    .doc-feedback { font-size: .8rem; margin-top: -.4rem; margin-bottom: .6rem; }

    /* ---------- Login ---------- */
    .login-panel { max-width: 470px; margin: 7vh auto; }
    .login-panel .card { box-shadow: 0 8px 30px rgba(44,62,80,.15); }

    /* ---------- Selo de modo demo ---------- */
    .modo-demo { position: fixed; top: 0; right: 0; z-index: 9999;
      background:#f39c12; color:#fff; padding:3px 12px; font-size:11px;
      font-weight:600; letter-spacing:.5px; border-bottom-left-radius:8px; }
  "))),
  uiOutput("tela")
)

# =============================================================================
#  SERVIDOR
# =============================================================================
server <- function(input, output, session) {

  # Conecta backend (no-op em DEMO) e garante que as abas existam na planilha
  backend_conectar()
  backend_inicializar()   # cria abas/cabeçalhos que ainda não existem (produção)

  # --- Estado da sessão -----------------------------------------------------
  auth      <- reactiveVal(FALSE)
  user_email<- reactiveVal(NULL)
  user_papel<- reactiveVal(NULL)

  v_plano   <- reactiveVal(db_plano_contas)
  v_diario  <- reactiveVal(schema_diario)
  v_saldos  <- reactiveVal(schema_saldos)
  v_users   <- reactiveVal(seed_usuarios)
  v_enc     <- reactiveVal(schema_encerramentos)
  aviso_enc <- reactiveVal(FALSE)   # controla o aviso de virada de ano (1x/sessão)

  # Carrega dados do backend ao iniciar (produção)
  if (!CFG$MODO_DEMO && nzchar(CFG$SHEET_ID)) {
    v_plano(  { p <- backend_ler(ABAS$plano, db_plano_contas); if (nrow(p)) p else db_plano_contas })
    v_diario( backend_ler(ABAS$diario, schema_diario))
    v_saldos( backend_ler(ABAS$saldos, schema_saldos))
    v_enc(    backend_ler(ABAS$encerramentos, schema_encerramentos))
    u <- backend_ler(ABAS$usuarios, schema_usuarios)
    v_users(if (nrow(u)) u else seed_usuarios)
  }

  # Helper de persistência (grava reactiveVal -> backend)
  persistir <- function(aba, df) backend_gravar(aba, df)

  # ==========================================================================
  #  TELA (login vs sistema)
  # ==========================================================================
  output$tela <- renderUI({
    if (!auth()) {
      tagList(
        if (CFG$MODO_DEMO) tags$div(class = "modo-demo", "MODO DEMO (sem Google)"),
        div(class = "login-panel",
          card(
            card_header(class = "bg-primary text-white text-center",
                        h4("Coordenação Financeira - PUC Jr", class = "mb-0")),
            card_body(
              if (CFG$MODO_DEMO) {
                tagList(
                  p(class = "text-muted",
                    "Modo de desenvolvimento. Informe um e-mail para simular o login."),
                  textInput("login_email", "E-mail (@gmail.com):",
                            value = CFG$ADMINS[1]),
                  helpText("E-mails de administrador entram como 'contador'."),
                  actionButton("btn_login", "Entrar (demo)",
                               class = "btn-primary w-100", icon = icon("right-to-bracket"))
                )
              } else {
                tagList(
                  p(class = "text-muted", "Acesse com seu e-mail e senha."),
                  textInput("login_email", "E-mail (@gmail.com):"),
                  passwordInput("login_senha", "Senha:"),
                  helpText("No primeiro acesso, a senha que você digitar será definida como sua senha."),
                  actionButton("btn_login", "Entrar", class = "btn-primary w-100",
                               icon = icon("right-to-bracket"))
                )
              }
            )
          )
        )
      )
    } else {
      contador <- identical(user_papel(), "contador")
      page_navbar(
        title = div(style = "display:flex;align-items:center;",
          tags$i(class = "fas fa-chart-line", style = "margin-right:10px;"),
          "Coordenação Financeira - PUC Jr"),
        id = "nav",

        # --- Diário (todos) ---
        nav_panel("Diário",
          layout_sidebar(
            sidebar = sidebar(title = "Novo lançamento", width = 380,
              dateInput("lan_data", "Data do fato:", value = Sys.Date()),
              radioButtons("lan_tipo", "Tipo de lançamento:",
                c("Simples (1 débito × 1 crédito)" = "simples",
                  "Composto (vários débitos/créditos)" = "composto"),
                selected = "simples"),

              # ---- Lançamento simples ----
              conditionalPanel("input.lan_tipo == 'simples'",
                selectizeInput("lan_debito", "Conta Débito:", choices = NULL),
                selectizeInput("lan_credito", "Conta Crédito:", choices = NULL),
                numericInput("lan_valor", "Valor (R$):", value = 0, min = 0.01, step = 0.01)
              ),

              # ---- Lançamento composto ----
              conditionalPanel("input.lan_tipo == 'composto'",
                div(class = "form-section", "Contas DEBITADAS"),
                do.call(tagList, lapply(1:5, function(i) fluidRow(
                  column(7, selectizeInput(paste0("lan_d_conta_", i), NULL, choices = NULL,
                                           options = list(placeholder = paste("Débito", i)))),
                  column(5, numericInput(paste0("lan_d_valor_", i), NULL, value = NA, min = 0, step = 0.01))
                ))),
                div(class = "form-section", "Contas CREDITADAS"),
                do.call(tagList, lapply(1:5, function(i) fluidRow(
                  column(7, selectizeInput(paste0("lan_c_conta_", i), NULL, choices = NULL,
                                           options = list(placeholder = paste("Crédito", i)))),
                  column(5, numericInput(paste0("lan_c_valor_", i), NULL, value = NA, min = 0, step = 0.01))
                ))),
                uiOutput("lan_comp_saldo")
              ),

              div(class = "form-section", "Identificação do terceiro (opcional)"),
              textInput("lan_razao", "Razão Social / Nome:",
                        placeholder = "ex.: ACME Serviços Ltda"),
              textInput("lan_docid", "CNPJ / CPF:",
                        placeholder = "somente números ou formatado"),
              uiOutput("lan_docid_feedback"),
              div(class = "form-section", "Histórico"),
              textAreaInput("lan_hist", "Descrição do fato contábil:", rows = 3),
              fileInput("lan_doc", "Comprovantes de suporte:",
                        multiple = TRUE, buttonLabel = "Selecionar",
                        placeholder = "até 5 arquivos, qualquer formato"),
              actionButton("btn_lanc", "Salvar lançamento",
                           class = "btn-success w-100", icon = icon("floppy-disk")),
              hr(),
              helpText("Lançamentos não podem ser editados/excluídos. ",
                       "Para corrigir, use o botão Estornar na tabela.")
            ),
            card(card_header("Livro Diário (cumulativo)"), DTOutput("tab_diario"))
          )
        ),

        # --- Demonstrações ---
        nav_panel("Demonstrações",
          navset_card_pill(
            nav_panel("Balanço Patrimonial",
              card(card_header("Balanço Patrimonial — posição atual"),
                   tableOutput("tab_balanco"))),
            nav_panel("DRE",
              card(card_header(div(
                "Demonstração do Resultado",
                selectInput("dre_exerc", NULL, choices = NULL, width = "150px"))),
                tableOutput("tab_dre"))),
            nav_panel("Razão",
              card(card_header("Razão por conta"),
                selectizeInput("razao_conta", "Conta:", choices = NULL),
                DTOutput("tab_razao"))),
            nav_panel("Pagamentos por Credor",
              card(card_header(div(
                "Valores pagos por CNPJ/CPF (saídas de Caixa/Banco)",
                selectInput("rel_pag_ano", NULL, choices = NULL, width = "150px"))),
                helpText("Lançamentos com crédito em conta de Caixa e Equivalentes (1.1.1), agrupados pelo CNPJ/CPF informado no histórico."),
                DTOutput("tab_pagamentos"))),
            nav_panel("Recebimentos por Pagador",
              card(card_header(div(
                "Valores recebidos por CNPJ/CPF (entradas em Caixa/Banco)",
                selectInput("rel_rec_ano", NULL, choices = NULL, width = "150px"))),
                helpText("Lançamentos com débito em conta de Caixa e Equivalentes (1.1.1), agrupados pelo CNPJ/CPF informado no histórico."),
                DTOutput("tab_recebimentos")))
          )
        ),

        # --- Configurações (somente contador) ---
        if (contador) nav_menu("Configurações",
          nav_panel("Carga Inicial de Saldos", cargaSaldosUI("carga")),
          nav_panel("Plano de Contas",
            layout_sidebar(
              sidebar = sidebar(title = "Gerenciar conta",
                textInput("pc_cod", "Código:", placeholder = "ex: 1.1.1.01.XX"),
                textInput("pc_nome", "Nome:"),
                textInput("pc_sub", "Subgrupo:"),
                textInput("pc_g2", "Grupo Nível 2:"),
                selectInput("pc_grupo", "Grupo:",
                  c("Ativo", "Passivo", "Patrimônio Líquido", "Despesas",
                    "Receitas", "Variações Patrimoniais")),
                selectInput("pc_tipo", "Tipo:", c("Patrimonial", "Resultado")),
                actionButton("btn_conta", "Inserir / Alterar",
                             class = "btn-warning w-100")),
              card(card_header("Estrutura do Plano de Contas"), DTOutput("tab_plano"))
            )),
          nav_panel("Usuários",
            layout_sidebar(
              sidebar = sidebar(title = "Cadastrar usuário",
                textInput("u_email", "E-mail (@gmail.com):"),
                textInput("u_nome", "Nome:"),
                selectInput("u_papel", "Papel:", c("usuario", "contador")),
                passwordInput("u_senha", "Senha (opcional):"),
                helpText("Se deixar a senha em branco, o usuário a define no primeiro acesso."),
                actionButton("btn_user", "Salvar usuário",
                             class = "btn-primary w-100")),
              card(card_header("Usuários cadastrados"), DTOutput("tab_users"))
            )),
          nav_panel("Encerramento Anual",
            layout_sidebar(
              sidebar = sidebar(title = "Encerrar exercício",
                selectInput("enc_exerc", "Exercício a encerrar:", choices = NULL),
                helpText("Apura o resultado e zera as contas de receita e ",
                         "despesa contra a conta Superávit/Déficit do Período ",
                         "(2.3.3), em 31/12. O resultado permanece em 2.3.3."),
                actionButton("btn_enc", "Executar encerramento (31/12)",
                             class = "btn-danger w-100", icon = icon("lock"))),
              card(card_header("Histórico de encerramentos"), DTOutput("tab_enc"))
            )),
          nav_panel("Backup / Exportação",
            card(
              card_header("Backup e exportação de dados"),
              p("Baixe os dados em Excel (.xlsx) para guardar nos computadores da PUC Jr."),
              selectInput("bk_exerc", "Exercício:", choices = NULL, width = "240px"),
              downloadButton("dl_exercicio",
                             "Baixar exercício selecionado (Diário + BP + DRE)",
                             class = "btn-primary"),
              br(), br(),
              downloadButton("dl_completo",
                             "Backup completo — todos os dados do sistema",
                             class = "btn-outline-secondary"),
              hr(),
              helpText("Recomendado: ao encerrar cada exercício, baixe o backup do ano e guarde uma cópia local. ",
                       "No encerramento, o sistema também gera no Drive uma planilha separada por exercício, ",
                       "o que mantém os dados históricos distribuídos em arquivos individuais.")
            ))
        ),

        nav_spacer(),
        nav_item(tags$span(class = "navbar-text me-2",
          textOutput("lbl_user", inline = TRUE))),
        nav_item(actionLink("btn_logout", "Sair", class = "text-danger",
                            icon = icon("right-from-bracket")))
      )
    }
  })

  # ==========================================================================
  #  LOGIN  (DEMO = só e-mail | Produção = e-mail + senha)
  # ==========================================================================
  resolver_papel <- function(email) {
    email <- tolower(trimws(email))
    if (email %in% tolower(CFG$ADMINS)) return("contador")
    u <- v_users() %>% filter(tolower(email) == !!email, ativo %in% c(TRUE, NA))
    if (nrow(u) > 0) return(u$papel[1])
    NA_character_
  }

  # Garante que um e-mail autorizado tenha uma linha em v_users (ex.: admins)
  garantir_usuario <- function(email, papel) {
    email <- tolower(trimws(email))
    df <- v_users()
    if (!email %in% tolower(df$email)) {
      df <- df %>% add_row(email = email, nome = email, papel = papel,
                           ativo = TRUE, senha_hash = "",
                           criado_em = as.character(Sys.time()))
      v_users(df); persistir("usuarios", df)
    }
    invisible(TRUE)
  }

  observeEvent(input$btn_login, {
    req(input$login_email)
    email <- tolower(trimws(input$login_email))
    papel <- resolver_papel(email)
    if (is.na(papel)) {
      showNotification("E-mail não autorizado. Peça ao contador para cadastrá-lo.",
                       type = "error"); return()
    }

    # DEMO: entra só com o e-mail
    if (CFG$MODO_DEMO) {
      user_email(email); user_papel(papel); auth(TRUE)
      showNotification(paste("Bem-vindo!", email), type = "message")
      return()
    }

    # PRODUÇÃO: e-mail + senha
    senha <- input$login_senha %||% ""
    if (!nzchar(trimws(senha))) { showNotification("Informe a senha.", type = "warning"); return() }
    garantir_usuario(email, papel)
    df <- v_users()
    hash_atual <- df$senha_hash[tolower(df$email) == email][1]
    hash_atual <- if (length(hash_atual) == 0 || is.na(hash_atual)) ""
                  else trimws(tolower(as.character(hash_atual)))
    if (!nzchar(hash_atual)) {
      # Primeiro acesso: define a senha digitada
      df$senha_hash[tolower(df$email) == email] <- hash_senha(senha)
      v_users(df); persistir("usuarios", df)
      user_email(email); user_papel(papel); auth(TRUE)
      showNotification("Senha definida e acesso liberado.", type = "message")
    } else if (isTRUE(hash_atual == hash_senha(senha))) {
      user_email(email); user_papel(papel); auth(TRUE)
      showNotification(paste("Bem-vindo!", email), type = "message")
    } else {
      showNotification("Senha incorreta.", type = "error")
    }
  })

  observeEvent(input$btn_logout, {
    auth(FALSE); user_email(NULL); user_papel(NULL)
    showNotification("Sessão encerrada.", type = "message")
  })

  output$lbl_user <- renderText({
    req(user_email()); paste0(user_email(), " (", user_papel(), ")")
  })

  # ==========================================================================
  #  ATUALIZAÇÃO DE SELECTS
  # ==========================================================================
  observe({
    req(auth()); p <- v_plano()
    if (nrow(p) > 0) {
      escolhas <- setNames(p$Codigo, paste(p$Codigo, "—", p$Nome))
      updateSelectizeInput(session, "lan_debito", choices = escolhas, server = TRUE)
      updateSelectizeInput(session, "lan_credito", choices = escolhas, server = TRUE)
      updateSelectizeInput(session, "razao_conta", choices = escolhas, server = TRUE)
      # Campos do lançamento composto (5 débitos + 5 créditos)
      for (i in 1:5) {
        updateSelectizeInput(session, paste0("lan_d_conta_", i),
                             choices = escolhas, server = TRUE)
        updateSelectizeInput(session, paste0("lan_c_conta_", i),
                             choices = escolhas, server = TRUE)
      }
    }
  })

  # Coleta as "pernas" (conta + valor) preenchidas do lançamento composto
  coletar_pernas <- function(prefixo) {
    pernas <- lapply(1:5, function(i) {
      conta <- input[[paste0(prefixo, "_conta_", i)]] %||% ""
      valor <- suppressWarnings(as.numeric(input[[paste0(prefixo, "_valor_", i)]]))
      if (nzchar(conta) && !is.na(valor) && valor > 0)
        tibble(Conta = conta, Valor = valor) else NULL
    })
    bind_rows(pernas)
  }

  # Saldo ao vivo do lançamento composto (confere débitos × créditos)
  output$lan_comp_saldo <- renderUI({
    deb <- coletar_pernas("lan_d"); cre <- coletar_pernas("lan_c")
    td <- if (nrow(deb)) sum(deb$Valor) else 0
    tc <- if (nrow(cre)) sum(cre$Valor) else 0
    bal <- abs(td - tc) < 0.005
    cor <- if (bal && td > 0) "alert-success" else "alert-warning"
    div(class = paste("alert mt-2", cor),
        tags$div(paste0("Total débitos: ", formatar_moeda(td))),
        tags$div(paste0("Total créditos: ", formatar_moeda(tc))),
        tags$div(class = "fw-bold",
          if (td == 0 && tc == 0) "Preencha as contas e valores."
          else if (bal) "Partida equilibrada ✓"
          else paste0("Diferença: ", formatar_moeda(abs(td - tc)),
                      " — ajuste para débitos = créditos.")))
  })
  observe({
    req(auth())
    anos <- v_diario()$Exercicio %>% unique() %>% na.omit() %>% sort(decreasing = TRUE)
    if (length(anos) == 0) anos <- exercicio_de(Sys.Date())
    updateSelectInput(session, "dre_exerc", choices = anos, selected = anos[1])
    ja <- v_enc()$Exercicio
    encerraveis <- setdiff(anos, ja)
    updateSelectInput(session, "enc_exerc",
                      choices = if (length(encerraveis)) encerraveis else anos)
    opc_ano <- c("Todos", as.character(anos))
    sel_ano <- as.character(exercicio_de(Sys.Date()))
    if (!sel_ano %in% opc_ano) sel_ano <- "Todos"
    updateSelectInput(session, "rel_pag_ano", choices = opc_ano, selected = sel_ano)
    updateSelectInput(session, "rel_rec_ano", choices = opc_ano, selected = sel_ano)
    updateSelectInput(session, "bk_exerc", choices = opc_ano, selected = sel_ano)
  })

  # --- Aviso de virada de ano: alerta o contador sobre exercícios pendentes --
  observe({
    req(auth(), identical(user_papel(), "contador"), !aviso_enc())
    d <- v_diario()
    if (nrow(d) == 0) return()
    ano_atual <- exercicio_de(Sys.Date())
    anos_mov <- d %>%
      filter(is.na(Tipo_Lancamento) | Tipo_Lancamento != "Encerramento") %>%
      pull(Exercicio) %>% unique() %>% na.omit()
    pendentes <- sort(setdiff(anos_mov[anos_mov < ano_atual], v_enc()$Exercicio))
    if (length(pendentes) > 0) {
      aviso_enc(TRUE)  # mostra apenas uma vez por sessão
      updateSelectInput(session, "enc_exerc", selected = as.character(pendentes[1]))
      showModal(modalDialog(
        title = tagList(icon("calendar-check"), " Encerramento pendente"),
        div(class = "alert alert-info",
          p(paste0("O ano corrente é ", ano_atual,
                   " e há exercício(s) anterior(es) ainda não encerrado(s): ",
                   paste(pendentes, collapse = ", "), ".")),
          p("Após revisar os lançamentos e ajustes, vá em ",
            tags$b("Configurações → Encerramento Anual"),
            " para apurar o resultado e gerar o arquivo do exercício."),
          p(class = "text-muted mb-0",
            "O encerramento permanece sendo uma ação confirmada por você.")),
        footer = modalButton("Entendi"), size = "m"))
    }
  })

  # ==========================================================================
  #  MÓDULO CARGA DE SALDOS
  # ==========================================================================
  cargaSaldosServer("carga", plano = v_plano, diario = v_diario,
                    saldos = v_saldos, persistir = persistir)

  # ==========================================================================
  #  NOVO LANÇAMENTO
  # ==========================================================================
  # Feedback ao vivo da validação de CNPJ/CPF
  output$lan_docid_feedback <- renderUI({
    info <- analisar_doc(input$lan_docid)
    if (info$status == "vazio") return(NULL)
    if (info$status == "valido")
      div(class = "doc-feedback text-success",
          icon("circle-check"), paste0(" ", info$tipo, " válido: ", info$fmt))
    else
      div(class = "doc-feedback text-danger",
          icon("circle-xmark"), " Documento inválido. Verifique os dígitos.")
  })

  observeEvent(input$btn_lanc, {
    tipo <- input$lan_tipo %||% "simples"

    # ---- Reúne as pernas conforme o tipo ----
    if (tipo == "simples") {
      req(input$lan_debito, input$lan_credito, input$lan_valor)
      v <- suppressWarnings(as.numeric(input$lan_valor))
      if (is.na(v) || v <= 0) { showNotification("Valor deve ser > 0.", type = "warning"); return() }
      if (input$lan_debito == input$lan_credito) {
        showNotification("Débito e crédito não podem ser a mesma conta.", type = "warning"); return() }
      pernas_d <- tibble(Conta = input$lan_debito, Valor = v)
      pernas_c <- tibble(Conta = input$lan_credito, Valor = v)
    } else {
      pernas_d <- coletar_pernas("lan_d")
      pernas_c <- coletar_pernas("lan_c")
      if (nrow(pernas_d) == 0 || nrow(pernas_c) == 0) {
        showNotification("Informe ao menos uma conta debitada e uma creditada (com valor > 0).",
                         type = "warning"); return() }
      if (abs(sum(pernas_d$Valor) - sum(pernas_c$Valor)) > 0.005) {
        showNotification(paste0("Partida desequilibrada: débitos = ",
          formatar_moeda(sum(pernas_d$Valor)), ", créditos = ",
          formatar_moeda(sum(pernas_c$Valor)), ". Devem ser iguais."),
          type = "error"); return() }
      if (length(intersect(pernas_d$Conta, pernas_c$Conta)) > 0) {
        showNotification("Uma mesma conta não pode estar debitada e creditada no mesmo lançamento.",
                         type = "warning"); return() }
    }

    # Validação do CNPJ/CPF (opcional, mas se preenchido deve ser válido)
    doc <- analisar_doc(input$lan_docid)
    if (doc$status == "invalido") {
      showNotification("CNPJ/CPF inválido. Corrija ou deixe em branco.",
                       type = "error"); return() }

    # Limite de anexos
    if (!is.null(input$lan_doc) && nrow(input$lan_doc) > 5) {
      showNotification("Anexe no máximo 5 arquivos por lançamento.",
                       type = "warning"); return() }

    # Composição do histórico: identificação do terceiro + texto digitado
    razao <- trimws(input$lan_razao %||% "")
    ident <- character(0)
    if (nzchar(razao))               ident <- c(ident, razao)
    if (doc$status == "valido")      ident <- c(ident, paste0(doc$tipo, ": ", doc$fmt))
    ident_txt <- paste(ident, collapse = " — ")
    hist_txt  <- trimws(input$lan_hist %||% "")
    historico_final <- if (nzchar(ident_txt) && nzchar(hist_txt))
                          paste0(ident_txt, " | ", hist_txt)
                       else if (nzchar(ident_txt)) ident_txt
                       else hist_txt

    link <- "Nenhum documento anexado"
    if (!is.null(input$lan_doc) && nrow(input$lan_doc) > 0) {
      links <- vapply(seq_len(nrow(input$lan_doc)), function(i)
        backend_upload_doc(input$lan_doc$datapath[i], input$lan_doc$name[i]),
        character(1))
      link <- paste(links, collapse = " | ")
      showNotification(paste(nrow(input$lan_doc),
                             "comprovante(s) processado(s)."),
                       type = "message", duration = 2)
    }

    # ---- Monta as linhas, agrupadas por Lancamento_ID ----
    base_id <- if (nrow(v_diario()) == 0) 1 else max(v_diario()$ID, na.rm = TRUE) + 1
    comum <- list(Historico = historico_final, Doc_Link = link,
                  Tipo_Lancamento = "Movimento", Ref_ID = NA_real_,
                  Exercicio = exercicio_de(input$lan_data),
                  Usuario = user_email() %||% "n/d",
                  Timestamp = as.character(Sys.time()),
                  Data = as.Date(input$lan_data))

    if (nrow(pernas_d) == 1 && nrow(pernas_c) == 1) {
      # 1×1: uma única linha com débito e crédito reais (contrapartida no razão)
      novo <- tibble(
        ID = base_id, Lancamento_ID = base_id, Data = comum$Data,
        Conta_Debito = pernas_d$Conta[1], Conta_Credito = pernas_c$Conta[1],
        Valor = pernas_d$Valor[1], Historico = comum$Historico, Doc_Link = comum$Doc_Link,
        Tipo_Lancamento = comum$Tipo_Lancamento, Ref_ID = comum$Ref_ID,
        Exercicio = comum$Exercicio, Usuario = comum$Usuario, Timestamp = comum$Timestamp)
    } else {
      # Composto: uma "perna" por conta (débito OU crédito), mesmo Lancamento_ID
      pernas <- bind_rows(
        transmute(pernas_d, Debito = Conta, Credito = "", Valor),
        transmute(pernas_c, Debito = "", Credito = Conta, Valor))
      novo <- tibble(
        ID = base_id + seq_len(nrow(pernas)) - 1, Lancamento_ID = base_id,
        Data = comum$Data,
        Conta_Debito = pernas$Debito, Conta_Credito = pernas$Credito,
        Valor = pernas$Valor, Historico = comum$Historico, Doc_Link = comum$Doc_Link,
        Tipo_Lancamento = comum$Tipo_Lancamento, Ref_ID = comum$Ref_ID,
        Exercicio = comum$Exercicio, Usuario = comum$Usuario, Timestamp = comum$Timestamp)
    }
    d <- bind_rows(v_diario(), novo); v_diario(d); persistir("diario", d)
    showNotification(if (nrow(pernas_d) == 1 && nrow(pernas_c) == 1) "Lançamento registrado!"
                     else paste0("Lançamento composto registrado (",
                                 nrow(pernas_d), " débito(s), ",
                                 nrow(pernas_c), " crédito(s))."),
                     type = "message")

    # Limpa os campos
    updateNumericInput(session, "lan_valor", value = 0)
    updateTextAreaInput(session, "lan_hist", value = "")
    updateTextInput(session, "lan_razao", value = "")
    updateTextInput(session, "lan_docid", value = "")
    for (i in 1:5) {
      updateNumericInput(session, paste0("lan_d_valor_", i), value = NA)
      updateNumericInput(session, paste0("lan_c_valor_", i), value = NA)
      updateSelectizeInput(session, paste0("lan_d_conta_", i), selected = "")
      updateSelectizeInput(session, paste0("lan_c_conta_", i), selected = "")
    }
  })

  # --- ESTORNO (única forma de "correção") ---
  observeEvent(input$estornar, {
    gid <- as.numeric(input$estornar)
    grupo <- v_diario() %>% mutate(LID = dplyr::coalesce(Lancamento_ID, ID)) %>%
      filter(LID == gid)
    if (nrow(grupo) == 0) return()
    if (grupo$Tipo_Lancamento[1] == "Estorno") {
      showNotification("Não é possível estornar um estorno.", type = "warning"); return() }
    if (gid %in% (v_diario()$Ref_ID %>% na.omit())) {
      showNotification("Lançamento já estornado.", type = "warning"); return() }
    n_pernas <- nrow(grupo)
    showModal(modalDialog(title = "Confirmar estorno",
      p(paste0("Estornar o lançamento #", gid,
               if (n_pernas > 1) paste0(" (composto, ", n_pernas, " partidas)") else "", "?")),
      p(class = "text-muted", "Será criado um lançamento inverso. O original é preservado."),
      footer = tagList(modalButton("Cancelar"),
        actionButton("conf_estorno", "Confirmar estorno", class = "btn-danger"))))
  })
  observeEvent(input$conf_estorno, {
    removeModal()
    gid <- as.numeric(input$estornar)
    grupo <- v_diario() %>% mutate(LID = dplyr::coalesce(Lancamento_ID, ID)) %>%
      filter(LID == gid)
    if (nrow(grupo) == 0) return()
    base_id <- max(v_diario()$ID, na.rm = TRUE) + 1
    est <- tibble(
      ID = base_id + seq_len(nrow(grupo)) - 1,
      Lancamento_ID = base_id,
      Data = Sys.Date(),
      Conta_Debito = grupo$Conta_Credito,   # inverte os lados
      Conta_Credito = grupo$Conta_Debito,
      Valor = grupo$Valor,
      Historico = paste0("ESTORNO do lanç. #", gid, " — ", grupo$Historico),
      Doc_Link = grupo$Doc_Link, Tipo_Lancamento = "Estorno", Ref_ID = gid,
      Exercicio = exercicio_de(Sys.Date()),
      Usuario = user_email() %||% "n/d", Timestamp = as.character(Sys.time())
    )
    d <- bind_rows(v_diario(), est); v_diario(d); persistir("diario", d)
    showNotification(paste0("Lançamento #", gid, " estornado."), type = "message")
  })

  # ==========================================================================
  #  TABELA DO DIÁRIO (uma linha por lançamento; compostos são agrupados)
  # ==========================================================================
  output$tab_diario <- renderDT({
    d <- v_diario()
    if (nrow(d) == 0)
      return(datatable(tibble(Mensagem = "Nenhum lançamento"), options = list(dom = "t"), rownames = FALSE))
    estornados <- d$Ref_ID %>% na.omit() %>% unique()
    nao_vazio <- function(x) x[!is.na(x) & x != ""]
    g <- d %>%
      mutate(LID = dplyr::coalesce(Lancamento_ID, ID),
             deb_amt = ifelse(!is.na(Conta_Debito) & Conta_Debito != "", Valor, 0)) %>%
      group_by(LID) %>%
      summarise(
        Data = dplyr::first(Data),
        Débito  = paste(unique(nao_vazio(Conta_Debito)),  collapse = "; "),
        Crédito = paste(unique(nao_vazio(Conta_Credito)), collapse = "; "),
        ValorN  = sum(deb_amt, na.rm = TRUE),
        Histórico = dplyr::first(Historico),
        Tipo = dplyr::first(Tipo_Lancamento),
        Documento = dplyr::first(Doc_Link),
        .groups = "drop") %>%
      arrange(desc(LID)) %>%
      mutate(
        Ação = ifelse(Tipo == "Estorno" | LID %in% estornados, "—",
          sprintf("<button class='btn btn-sm btn-outline-danger' onclick='Shiny.setInputValue(\"estornar\", %d, {priority:\"event\"})'>Estornar</button>", LID)),
        Valor = formatar_moeda(ValorN), Data = format(Data, "%d/%m/%Y")) %>%
      select(`Nº` = LID, Data, Débito, Crédito, Valor,
             Histórico, Tipo, Documento, Ação)
    datatable(g, escape = FALSE, rownames = FALSE,
              options = list(pageLength = 12, scrollX = TRUE))
  })

  # ==========================================================================
  #  PLANO DE CONTAS
  # ==========================================================================
  observeEvent(input$btn_conta, {
    req(input$pc_cod, input$pc_nome)
    df <- v_plano()
    if (input$pc_cod %in% df$Codigo) {
      df <- df %>% mutate(
        Nome = if_else(Codigo == input$pc_cod, input$pc_nome, Nome),
        Subgrupo = if_else(Codigo == input$pc_cod, input$pc_sub, Subgrupo),
        Grupo_Nivel2 = if_else(Codigo == input$pc_cod, input$pc_g2, Grupo_Nivel2),
        Grupo = if_else(Codigo == input$pc_cod, input$pc_grupo, Grupo),
        Tipo = if_else(Codigo == input$pc_cod, input$pc_tipo, Tipo))
      showNotification("Conta alterada.", type = "message")
    } else {
      df <- df %>% add_row(Codigo = input$pc_cod, Nome = input$pc_nome,
        Subgrupo = input$pc_sub, Grupo_Nivel2 = input$pc_g2,
        Grupo = input$pc_grupo, Tipo = input$pc_tipo,
        Nivel = str_count(input$pc_cod, "\\.") + 1)
      showNotification("Conta inserida.", type = "message")
    }
    df <- df %>% arrange(Codigo); v_plano(df); persistir("plano_contas", df)
    updateTextInput(session, "pc_cod", value = ""); updateTextInput(session, "pc_nome", value = "")
  })
  output$tab_plano <- renderDT({
    d <- v_diario()
    contas_com_mov <- unique(c(d$Conta_Debito, d$Conta_Credito))
    df <- v_plano() %>% select(Codigo, Nome, Subgrupo, Grupo, Tipo) %>%
      mutate(Ação = case_when(
        Codigo %in% CONTAS_SISTEMA ~
          "<span class='text-muted'>conta do sistema</span>",
        Codigo %in% contas_com_mov ~
          "<span class='text-muted'>possui lançamentos</span>",
        TRUE ~ sprintf("<button class='btn btn-sm btn-outline-danger' onclick='Shiny.setInputValue(\"excluir_conta\", \"%s\", {priority:\"event\"})'>Excluir</button>", Codigo)
      ))
    datatable(df, escape = FALSE,
              options = list(pageLength = 15, scrollX = TRUE), rownames = FALSE)
  })

  observeEvent(input$excluir_conta, {
    cod <- input$excluir_conta
    if (cod %in% CONTAS_SISTEMA) {
      showNotification("Esta conta é usada pelo sistema e não pode ser excluída.",
                       type = "warning"); return() }
    d <- v_diario()
    if (cod %in% c(d$Conta_Debito, d$Conta_Credito)) {
      showNotification(paste0("A conta ", cod,
        " possui lançamentos e não pode ser excluída. Estorne os lançamentos primeiro."),
        type = "warning", duration = 5); return() }
    nome <- v_plano() %>% filter(Codigo == cod) %>% pull(Nome)
    showModal(modalDialog(title = "Confirmar exclusão de conta",
      p(paste0("Excluir a conta ", cod, " — ", nome[1], "?")),
      p(class = "text-muted", "Só é possível excluir contas sem lançamentos."),
      footer = tagList(modalButton("Cancelar"),
        actionButton("conf_excluir_conta", "Excluir conta", class = "btn-danger"))))
  })

  observeEvent(input$conf_excluir_conta, {
    removeModal()
    cod <- input$excluir_conta
    if (cod %in% CONTAS_SISTEMA) return()
    d <- v_diario()
    if (cod %in% c(d$Conta_Debito, d$Conta_Credito)) return()
    df <- v_plano() %>% filter(Codigo != cod)
    v_plano(df); persistir("plano_contas", df)
    showNotification(paste0("Conta ", cod, " excluída."), type = "message")
  })

  # ==========================================================================
  #  USUÁRIOS
  # ==========================================================================
  observeEvent(input$btn_user, {
    req(input$u_email, input$u_nome)
    em <- tolower(trimws(input$u_email)); df <- v_users()
    nova_senha <- input$u_senha %||% ""
    if (em %in% tolower(df$email)) {
      df <- df %>% mutate(
        nome = if_else(tolower(email) == em, input$u_nome, nome),
        papel = if_else(tolower(email) == em, input$u_papel, papel),
        senha_hash = if_else(tolower(email) == em & nzchar(nova_senha),
                             hash_senha(nova_senha), senha_hash))
      showNotification(if (nzchar(nova_senha))
        "Usuário atualizado (senha redefinida)." else "Usuário atualizado.",
        type = "message")
    } else {
      df <- df %>% add_row(email = em, nome = input$u_nome, papel = input$u_papel,
                           ativo = TRUE,
                           senha_hash = if (nzchar(nova_senha)) hash_senha(nova_senha) else "",
                           criado_em = as.character(Sys.time()))
      showNotification("Usuário cadastrado.", type = "message")
    }
    v_users(df); persistir("usuarios", df)
    updateTextInput(session, "u_email", value = ""); updateTextInput(session, "u_nome", value = "")
    updateTextInput(session, "u_senha", value = "")
  })

  output$tab_users <- renderDT({
    adm <- tolower(CFG$ADMINS)
    df <- v_users() %>%
      mutate(Senha = ifelse(is.na(senha_hash) | !nzchar(senha_hash),
                            "não definida", "definida")) %>%
      select(-any_of("senha_hash")) %>%
      mutate(Ação = ifelse(
        tolower(email) %in% adm,
        "<span class='text-muted'>desenvolvedor (protegido)</span>",
        sprintf("<button class='btn btn-sm btn-outline-danger' onclick='Shiny.setInputValue(\"excluir_user\", \"%s\", {priority:\"event\"})'>Excluir</button>", email)))
    datatable(df, escape = FALSE, options = list(pageLength = 10), rownames = FALSE)
  })

  observeEvent(input$excluir_user, {
    em <- tolower(input$excluir_user)
    if (em %in% tolower(CFG$ADMINS)) {
      showNotification("Os desenvolvedores não podem ser excluídos.",
                       type = "warning"); return() }
    if (!is.null(user_email()) && em == tolower(user_email())) {
      showNotification("Você não pode excluir o usuário que está em uso.",
                       type = "warning"); return() }
    showModal(modalDialog(title = "Confirmar exclusão",
      p(paste0("Excluir o usuário ", em, "?")),
      p(class = "text-muted", "Esta ação remove o acesso do usuário ao sistema."),
      footer = tagList(modalButton("Cancelar"),
        actionButton("conf_excluir_user", "Excluir", class = "btn-danger"))))
  })

  observeEvent(input$conf_excluir_user, {
    removeModal()
    em <- tolower(input$excluir_user)
    if (em %in% tolower(CFG$ADMINS)) return()
    df <- v_users() %>% filter(tolower(email) != em)
    v_users(df); persistir("usuarios", df)
    showNotification("Usuário excluído.", type = "message")
  })

  # ==========================================================================
  #  DEMONSTRAÇÕES
  # ==========================================================================
  output$tab_balanco <- renderTable({
    d <- v_diario()
    if (nrow(d) == 0) return(tibble(Mensagem = "Sem lançamentos"))
    bp <- dados_balanco(d, v_plano())   # já inclui o resultado do período em 2.3.3
    if (nrow(bp) == 0) return(tibble(Mensagem = "Sem movimentação patrimonial"))

    bloco <- function(grupo_label) {
      sub <- bp %>% filter(Grupo == grupo_label)
      if (nrow(sub) == 0) return(tibble())
      total <- sum(sub$Saldo)
      linhas <- sub %>% arrange(Codigo) %>%
        transmute(Conta = paste0("   ", Codigo, " ", Conta),
                  Valor = formatar_moeda(Saldo))
      bind_rows(tibble(Conta = grupo_label, Valor = ""), linhas,
                tibble(Conta = paste("TOTAL", grupo_label), Valor = formatar_moeda(total)))
    }
    ativo   <- bloco("ATIVO")
    passivo <- bloco("PASSIVO + PATRIMÔNIO LÍQUIDO")
    bind_rows(ativo, tibble(Conta = "", Valor = ""), passivo)
  }, striped = TRUE, hover = TRUE, width = "100%", na = "")

  output$tab_dre <- renderTable({
    req(input$dre_exerc)
    ano <- as.integer(input$dre_exerc)
    d <- v_diario() %>% filter(Exercicio == ano,
                               is.na(Tipo_Lancamento) | Tipo_Lancamento != "Encerramento")
    if (nrow(d) == 0) return(tibble(Mensagem = "Sem lançamentos no exercício"))
    s <- saldos_por_conta(d, v_plano())
    rec <- s %>% filter(str_starts(Codigo, "4"), abs(Saldo) > 0.01)
    desp <- s %>% filter(str_starts(Codigo, "3"), abs(Saldo) > 0.01)
    tr <- sum(rec$Saldo); td <- sum(desp$Saldo); res <- tr - td

    linhas <- function(df) df %>% arrange(Codigo) %>%
      transmute(Conta = paste0("   ", Codigo, " ", Nome), Valor = formatar_moeda(Saldo))
    bind_rows(
      tibble(Conta = "RECEITAS", Valor = ""), linhas(rec),
      tibble(Conta = "(=) Total de Receitas", Valor = formatar_moeda(tr)),
      tibble(Conta = "", Valor = ""),
      tibble(Conta = "DESPESAS", Valor = ""), linhas(desp),
      tibble(Conta = "(=) Total de Despesas", Valor = formatar_moeda(td)),
      tibble(Conta = "", Valor = ""),
      tibble(Conta = "RESULTADO DO PERÍODO (Superávit/Déficit)",
             Valor = formatar_moeda(res))
    )
  }, striped = TRUE, hover = TRUE, width = "100%", na = "")

  output$tab_razao <- renderDT({
    req(input$razao_conta); d <- v_diario()
    if (nrow(d) == 0) return(datatable(tibble(Mensagem = "Sem lançamentos"),
                                       options = list(dom = "t"), rownames = FALSE))
    c <- input$razao_conta
    deb <- d %>% filter(Conta_Debito == c) %>%
      transmute(ID, Data, Tipo = "Débito",
                Contrapartida = as.character(Conta_Credito), Valor, Historico)
    cre <- d %>% filter(Conta_Credito == c) %>%
      transmute(ID, Data, Tipo = "Crédito",
                Contrapartida = as.character(Conta_Debito), Valor, Historico)
    r <- bind_rows(deb, cre)
    if (nrow(r) == 0)
      return(datatable(tibble(Mensagem = "Sem movimento nesta conta"),
                       options = list(dom = "t"), rownames = FALSE))
    r <- r %>%
      mutate(Contrapartida = ifelse(is.na(Contrapartida) | Contrapartida == "",
                                    "(composto)", Contrapartida)) %>%
      arrange(Data, ID) %>%
      mutate(Sinal = if_else(
               (str_starts(c, "1") | str_starts(c, "3")) == (Tipo == "Débito"),
               Valor, -Valor),
             Saldo = cumsum(Sinal),
             Data = format(Data, "%d/%m/%Y"),
             Valor = formatar_moeda(Valor), Saldo = formatar_moeda(Saldo)) %>%
      select(ID, Data, Tipo, Contrapartida, Valor, `Saldo acumulado` = Saldo, Historico)
    if (nrow(r) == 0) r <- tibble(Mensagem = "Sem movimento nesta conta")
    datatable(r, options = list(scrollX = TRUE, pageLength = 12), rownames = FALSE)
  })

  # --- Relatórios por CNPJ/CPF (caixa/banco) -------------------------------
  render_relatorio_caixa <- function(lado, ano) {
    r <- relatorio_caixa_terceiros(v_diario(), v_plano(), lado, ano)
    msg <- if (lado == "pagamento") "Sem pagamentos no período"
           else "Sem recebimentos no período"
    if (nrow(r) == 0)
      return(datatable(tibble(Mensagem = msg), options = list(dom = "t"), rownames = FALSE))
    total <- sum(r$Total, na.rm = TRUE)
    r <- r %>%
      transmute(`CNPJ/CPF` = Documento, `Razão Social / Nome` = Nome,
                `Nº lançamentos`, Total = formatar_moeda(Total)) %>%
      bind_rows(tibble(`CNPJ/CPF` = "TOTAL GERAL", `Razão Social / Nome` = "",
                       `Nº lançamentos` = sum(r$`Nº lançamentos`),
                       Total = formatar_moeda(total)))
    datatable(r, rownames = FALSE,
              options = list(pageLength = 15, scrollX = TRUE,
                             dom = "Bfrtip", buttons = c("copy", "csv", "excel")),
              extensions = "Buttons")
  }

  output$tab_pagamentos   <- renderDT(render_relatorio_caixa("pagamento", input$rel_pag_ano))
  output$tab_recebimentos <- renderDT(render_relatorio_caixa("recebimento", input$rel_rec_ano))

  # ==========================================================================
  #  ENCERRAMENTO ANUAL (31/12)
  # ==========================================================================
  output$tab_enc <- renderDT({
    e <- v_enc()
    if (nrow(e) == 0) return(datatable(tibble(Mensagem = "Nenhum encerramento"),
                                       options = list(dom = "t"), rownames = FALSE))
    e %>% mutate(across(c(Total_Receitas, Total_Despesas, Resultado), formatar_moeda),
                 Data_Encerramento = format(as.Date(Data_Encerramento), "%d/%m/%Y"),
                 Arquivo = if_else(
                   !is.na(Arquivo_Link) & str_starts(Arquivo_Link %||% "", "http"),
                   paste0("<a href='", Arquivo_Link, "' target='_blank'>Abrir planilha</a>"),
                   Arquivo_Link)) %>%
      select(Exercicio, Data_Encerramento, Total_Receitas, Total_Despesas,
             Resultado, Arquivo, Usuario) %>%
      datatable(escape = FALSE, options = list(pageLength = 10), rownames = FALSE)
  })

  # --- Backup / Exportação local (.xlsx) -----------------------------------
  output$dl_exercicio <- downloadHandler(
    filename = function() {
      a <- input$bk_exerc %||% "Todos"
      if (identical(a, "Todos")) paste0("PUCJr_diario_completo_", Sys.Date(), ".xlsx")
      else paste0("PUCJr_exercicio_", a, ".xlsx")
    },
    content = function(file) {
      a <- input$bk_exerc %||% "Todos"
      d <- v_diario()
      if (!identical(a, "Todos"))
        d <- d %>% filter(Exercicio == suppressWarnings(as.integer(a)))
      diario_tab <- d %>% arrange(ID) %>%
        mutate(Data = format(as.Date(Data), "%d/%m/%Y"))
      abas <- list(Diario = diario_tab)
      if (!identical(a, "Todos")) {
        ano <- as.integer(a)
        abas[["Balanco_Patrimonial"]] <- dados_balanco(v_diario(), v_plano(), ate_ano = ano)
        abas[["DRE"]] <- dados_dre(v_diario(), v_plano(), ano)
      } else {
        abas[["Balanco_Patrimonial"]] <- dados_balanco(v_diario(), v_plano())
      }
      writexl::write_xlsx(abas, path = file)
    }
  )

  output$dl_completo <- downloadHandler(
    filename = function() paste0("PUCJr_backup_completo_", Sys.Date(), ".xlsx"),
    content = function(file) {
      writexl::write_xlsx(list(
        plano_contas   = v_plano(),
        diario         = v_diario() %>% mutate(Data = format(as.Date(Data), "%d/%m/%Y")),
        saldos_iniciais = v_saldos() %>% mutate(Data_Base = as.character(Data_Base)),
        usuarios       = v_users() %>% select(-any_of("senha_hash")),
        encerramentos  = v_enc() %>% mutate(Data_Encerramento = as.character(Data_Encerramento))
      ), path = file)
    }
  )

  observeEvent(input$btn_enc, {
    req(input$enc_exerc)
    ano <- as.integer(input$enc_exerc)
    if (ano %in% v_enc()$Exercicio) {
      showNotification("Exercício já encerrado.", type = "warning"); return() }
    showModal(modalDialog(title = "Confirmar encerramento",
      p(paste0("Encerrar o exercício ", ano, " em 31/12/", ano, "?")),
      p(class = "text-muted",
        "As contas de receita e despesa serão zeradas contra a conta Superávit/Déficit do Período (2.3.3), onde o resultado permanecerá acumulado. Esta operação gera lançamentos definitivos."),
      footer = tagList(modalButton("Cancelar"),
        actionButton("conf_enc", "Confirmar encerramento", class = "btn-danger"))))
  })

  observeEvent(input$conf_enc, {
    removeModal()
    ano <- as.integer(input$enc_exerc)
    data_enc <- as.Date(paste0(ano, "-12-31"))
    CONTA_RES <- "2.3.3.XX.XX"  # Superávit ou Déficit do Período (resultado fica aqui)

    d_ano <- v_diario() %>% filter(Exercicio == ano,
              is.na(Tipo_Lancamento) | Tipo_Lancamento != "Encerramento")
    s <- saldos_por_conta(d_ano, v_plano())
    rec  <- s %>% filter(str_starts(Codigo, "4"), abs(Saldo) > 0.01)
    desp <- s %>% filter(str_starts(Codigo, "3"), abs(Saldo) > 0.01)
    if (nrow(rec) == 0 && nrow(desp) == 0) {
      showNotification("Sem contas de resultado a encerrar neste exercício.", type = "warning"); return() }

    novos <- tibble()
    # Zera receitas (saldo credor): Débito receita / Crédito 2.3.3
    for (i in seq_len(nrow(rec)))
      novos <- bind_rows(novos, tibble(Conta_Debito = rec$Codigo[i],
        Conta_Credito = CONTA_RES, Valor = abs(rec$Saldo[i]),
        Historico = paste("Encerramento", ano, "- apuração de receita em 2.3.3")))
    # Zera despesas (saldo devedor): Débito 2.3.3 / Crédito despesa
    for (i in seq_len(nrow(desp)))
      novos <- bind_rows(novos, tibble(Conta_Debito = CONTA_RES,
        Conta_Credito = desp$Codigo[i], Valor = abs(desp$Saldo[i]),
        Historico = paste("Encerramento", ano, "- apuração de despesa em 2.3.3")))

    tr <- sum(rec$Saldo); td <- sum(desp$Saldo); res <- tr - td
    # O resultado apurado PERMANECE em 2.3.3 (Superávit/Déficit do Período).
    # Não há transferência para o Patrimônio Social (2.3.1).

    prox <- max(v_diario()$ID, na.rm = TRUE)
    novos <- novos %>% mutate(ID = prox + row_number(), Data = data_enc,
      Doc_Link = "Encerramento automático", Tipo_Lancamento = "Encerramento",
      Ref_ID = NA_real_, Exercicio = ano, Usuario = user_email() %||% "n/d",
      Timestamp = as.character(Sys.time())) %>% select(all_of(names(schema_diario)))

    d <- bind_rows(v_diario(), novos); v_diario(d); persistir("diario", d)

    # ---- Gera o arquivo arquivado do exercício (Diário + BP + DRE) ----------
    diario_ano <- d %>% filter(Exercicio == ano) %>% arrange(ID) %>%
      mutate(Data = format(Data, "%d/%m/%Y")) %>%
      select(ID, Data, Conta_Debito, Conta_Credito, Valor, Historico,
             Doc_Link, Tipo_Lancamento, Ref_ID, Usuario)
    bp  <- dados_balanco(d, v_plano(), ate_ano = ano)
    dre <- dados_dre(d, v_plano(), ano)
    dre_final <- bind_rows(
      dre,
      tibble(Categoria = "TOTAL", Codigo = "", Conta = "Total de Receitas", Saldo = tr),
      tibble(Categoria = "TOTAL", Codigo = "", Conta = "Total de Despesas", Saldo = td),
      tibble(Categoria = "RESULTADO", Codigo = "", Conta = "Superávit/Déficit do Período", Saldo = res)
    )
    abas_arquivo <- list()
    abas_arquivo[[paste0("Diario_", ano)]]            <- diario_ano
    abas_arquivo[["Balanco_Patrimonial"]]             <- bp
    abas_arquivo[[paste0("DRE_", ano)]]               <- dre_final

    showNotification("Gerando arquivo do exercício...", type = "message", duration = 2)
    link_arq <- backend_arquivar_exercicio(ano, abas_arquivo)

    log <- v_enc() %>% add_row(Exercicio = ano, Data_Encerramento = data_enc,
      Total_Receitas = tr, Total_Despesas = td, Resultado = res,
      Arquivo_Link = link_arq,
      Usuario = user_email() %||% "n/d", Timestamp = as.character(Sys.time()))
    v_enc(log); persistir("encerramentos", log)
    showNotification(paste0("Exercício ", ano, " encerrado. Resultado: ",
                            formatar_moeda(res)), type = "message", duration = 6)
  })
}

shinyApp(ui = ui, server = server)
