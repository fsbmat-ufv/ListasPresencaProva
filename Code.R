rm(list=ls())
cat("\014")
library(tidyverse)
library(data.table)
df <- fread("Matriculados.csv")
df$NomeCompleto <- paste(df$Nome, df$Sobrenome)
df <- df %>% select(`Número de identificação`, NomeCompleto, Grupos)
names(df) <- c("Matrícula", "Nome", "Turma")
df <- df %>%
  mutate(Matrícula = if_else(Nome == "ELISIANE TAÍRES GONÇALVES VARELA",
                             "ES111941", Matrícula))
df <- df %>% arrange(Matrícula)
str(df)
df <- df %>%
  mutate(Mat_num = as.integer(sub("ES", "", Matrícula))) %>%
  arrange(Mat_num) %>%
  select(-Mat_num)
df[, Obs := .I]
df <- df %>% select(Obs, Matrícula, Nome, Turma)
df$"Assinatura - P1 - 12/9" <- ""
df$"Assinatura - P2 - 31/10" <- ""
df$"Assinatura - P3 - 28/11" <- ""

sizes <- c(60, 60, 80, 113, 65, 60)
stopifnot(sum(sizes) == nrow(df))

# Limites cumulativos
ends   <- cumsum(sizes)
starts <- c(1, head(ends, -1) + 1)

mk_slice <- function(i) slice(df, starts[i]:ends[i])

df1 <- mk_slice(1)
df2 <- mk_slice(2)
df3 <- mk_slice(3)
df4 <- mk_slice(4)
df5 <- mk_slice(5)
df6 <- mk_slice(6)


library(openxlsx)

# Liste os dataframes, títulos e nomes de arquivos
dfs <- list(
  df1 = df1,
  df2 = df2,
  df3 = df3,
  df4 = df4,
  df5 = df5,
  df6 = df6
)

titulos <- c(
  "EST105 - II/2025 - PVA 153 – 60 estudantes",
  "EST105 - II/2025 - PVA 165 – 60 estudantes",
  "EST105 - II/2025 - PVA 179 – 80 estudantes",
  "EST105 - II/2025 - PVA 201 – 113 estudantes",
  "EST105 - II/2025 - PVA 277 – 65 estudantes",
  "EST105 - II/2025 - PVA 279 – 60 estudantes"
)

# Sugestão de nomes de arquivos (pode ajustar se quiser)
arquivos <- c(
  "df1_PVA153.xlsx",
  "df2_PVA165.xlsx",
  "df3_PVA179.xlsx",
  "df4_PVA201.xlsx",
  "df5_PVA277.xlsx",
  "df6_PVA279.xlsx"
)

# Estilo opcional para o título
estilo_titulo <- createStyle(halign = "center", textDecoration = "bold", fontSize = 12)

# Função utilitária para escrever um df em um arquivo com título mesclado
escreve_xlsx_com_titulo <- function(dado, titulo, arquivo) {
  wb <- createWorkbook()
  addWorksheet(wb, "Planilha")
  # escreve o df começando na linha 2
  writeData(wb, sheet = 1, x = dado, startRow = 2)
  # mescla a linha 1 em todas as colunas
  mergeCells(wb, sheet = 1, cols = 1:ncol(dado), rows = 1)
  # escreve o título
  writeData(wb, sheet = 1, x = titulo, startRow = 1, startCol = 1)
  # aplica estilo e ajustes (opcionais)
  addStyle(wb, sheet = 1, style = estilo_titulo, rows = 1, cols = 1)
  setRowHeights(wb, sheet = 1, rows = 1, heights = 22)
  setColWidths(wb, sheet = 1, cols = 1:ncol(dado), widths = "auto")
  # salva
  saveWorkbook(wb, arquivo, overwrite = TRUE)
}

# Loop sobre os seis bancos
for (i in seq_along(dfs)) {
  escreve_xlsx_com_titulo(dado = dfs[[i]], titulo = titulos[i], arquivo = arquivos[i])
}


#Exportar para o Word

library(officer)
library(flextable)

# seus dataframes e títulos
dfs <- list(df1, df2, df3, df4, df5, df6)
titulos <- c(
  "EST105 - II/2025 - PVA 153 – 60 estudantes",
  "EST105 - II/2025 - PVA 165 – 60 estudantes",
  "EST105 - II/2025 - PVA 179 – 80 estudantes",
  "EST105 - II/2025 - PVA 201 – 113 estudantes",
  "EST105 - II/2025 - PVA 277 – 65 estudantes",
  "EST105 - II/2025 - PVA 279 – 60 estudantes"
)
arquivos <- c(
  "df1_PVA153.docx",
  "df2_PVA165.docx",
  "df3_PVA179.docx",
  "df4_PVA201.docx",
  "df5_PVA277.docx",
  "df6_PVA279.docx"
)

# estilo de borda 1pt
borda_1pt <- fp_border(color = "black", width = 1)

gera_docx_paisagem <- function(dat, titulo, arquivo, max_width_in = 9.5) {
  # cria flextable
  ft <- flextable(dat)
  
  # linha de título mesclada
  ft <- add_header_row(ft, values = titulo, colwidths = ncol(dat))
  ft <- bold(ft, i = 1, part = "header")
  ft <- align(ft, i = 1, part = "header", align = "center")
  
  # centralizar texto em todas as células
  ft <- align(ft, align = "center", part = "all")
  ft <- valign(ft, valign = "center", part = "all")
  
  # bordas 1pt
  ft <- border_remove(ft)
  ft <- border_outer(ft, part = "all", border = borda_1pt)
  ft <- border_inner_h(ft, part = "all", border = borda_1pt)
  ft <- border_inner_v(ft, part = "all", border = borda_1pt)
  
  # ajuste de largura
  ft <- set_table_properties(ft, layout = "autofit")
  ft <- autofit(ft)
  ft <- fit_to_width(ft, max_width = max_width_in)
  
  # documento em paisagem com margens estreitas
  doc <- read_docx()
  sect <- prop_section(
    type = "continuous",
    page_size = page_size(orient = "landscape"),
    page_margins = page_mar(
      top = 0.5, bottom = 0.5, left = 0.5, right = 0.5, header = 0.3, footer = 0.3
    )
  )
  doc <- body_set_default_section(doc, value = sect)
  
  # adiciona a tabela
  doc <- body_add_flextable(doc, value = ft)
  print(doc, target = arquivo)
}

# gerar os 6 arquivos
for (i in seq_along(dfs)) {
  gera_docx_paisagem(dfs[[i]], titulos[i], arquivos[i])
}



