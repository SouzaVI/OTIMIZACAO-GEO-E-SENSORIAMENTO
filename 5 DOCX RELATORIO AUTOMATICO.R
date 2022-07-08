
#------------------------------------------------------------------------------
# Gerar relatorio em documento, o script le uma base dados e obtem informacoes de interesse
# Autor: Igor Viana Souza ( M.e Engenharia Florestal)
# Data de criacao: 01/05/2022
#-------------------------------------------------------------------------------

rm(list=ls(all=TRUE))

if(!require(officer)){install.packages("officer")}
if(!require(rgdal)){install.packages("rgdal")}
if(!require(tidyverse)){install.packages("tidyverse")}
if(!require(dplyr)){install.packages("dplyr")}
if(!require(sf)){install.packages("sf")}
if(!require(foreign)){install.packages("foreign")}
if(!require(readxl)){install.packages("readxl")}
if(!require(lubridate)){install.packages("lubridate")}
if(!require(magrittr)){install.packages("magrittr")}
if(!require(flextable)){install.packages("flextable")}
if(!require(plyr)){install.packages("plyr")}
if(!require(ggpubr)){install.packages("ggpubr")}
if(!require(scales)){install.packages("scales")}
if(!require(RColorBrewer)){install.packages("RColorBrewer")}
###############################################################################
#                           CONFIGURACAO PAGINA
###############################################################################
DATA_ATUAL <- "RESUMO DIÁRIO DE PREVISÃO DE AMOSTRAS - 07.07.2022"


#landscape section
landscape=block_section(
  prop_section(
    page_size = page_size(orient="landscape"), type = "nextPage",
    
  )
)

#portrait section
portrait=block_section(
  prop_section(
    page_size = page_size(orient="portrait"), type = "nextPage"
  )
)

text_style <- fp_text(font.size = 16)
text_style_2 <- fp_text(font.size = 16, bold = TRUE )
#cores <- brewer.pal(10, "RdYlGn")

###############################################################################  
#                              INICIO RESUMO
################################################################################

# Dados

df <- read_excel("G:/.shortcut-targets-by-id/1eY6qLaqCw0vMplswsEaSexf_xogKRn_4/GEOPROCESSAMENTO/RESUMO DIARIO - 2022/Controle Geral Amostras - 2022_v2_geo.xlsm", sheet = "BASE")

################################################################################
#                                 PENDENTES
################################################################################


pendentes_df<-filter(df, Tipo == "SOLO", Laboratório 
                     %in% c("EXATA", "IBRA", "SOLO E PLANTA"), Status != "FINALIZADO" )%>%
  select(Consultor,Laboratório, Lote, `Data Envio`, `Data Recebimento`, `Data de Entrada`,
         Cliente, Fazenda, `Nº Amostras`, `Previsão de Liberação`, `Data de Liberação`) 

pendentes_df$`Data Envio` <- as.Date(pendentes_df$`Data Envio`)
pendentes_df$`Data Recebimento` <- as.Date(pendentes_df$`Data Recebimento`)
pendentes_df$`Data de Entrada` <- as.Date(pendentes_df$`Data de Entrada`)
pendentes_df$`Previsão de Liberação` <- as.Date(pendentes_df$`Previsão de Liberação`)



pendentes_df<-pendentes_df[order(pendentes_df$`Data Envio`), na.last=FALSE]



#EXATA
pendentes_exata_df<- filter(pendentes_df, Laboratório == "EXATA")%>%
  select(Lote, Cliente, Fazenda, `Nº Amostras`, 
         `Data Envio`, `Data Recebimento`, `Data de Entrada`,
         `Previsão de Liberação`)  

pendentes_exata_total <- count(pendentes_exata_df$Fazenda)
pendentes_exata_total2<-sum(pendentes_exata_total$freq)
total_exata_lote<-paste0( "Nº de lotes pendentes: ", pendentes_exata_total2)

pendentes_exata_total3<-sum(pendentes_exata_df$`Nº Amostras`)
total_exata_amostra<-paste0( "Nº de amostras pendentes: ", pendentes_exata_total3)


#IBRA
pendentes_ibra_df<- filter(pendentes_df, Laboratório == "IBRA")%>%
  select(Lote, Cliente, Fazenda, `Nº Amostras`, 
         `Data Envio`, `Data Recebimento`,
         `Previsão de Liberação`)  

pendentes_ibra_total <- count(pendentes_ibra_df$Fazenda)
pendentes_ibra_total2<-sum(pendentes_ibra_total$freq)
total_ibra_lote<-paste0( "Nº de lotes pendentes: ", pendentes_ibra_total2)

pendentes_ibra_total3<-sum(pendentes_ibra_df$`Nº Amostras`)
total_ibra_amostra<-paste0( "Nº de amostras pendentes: ", pendentes_ibra_total3)



#SOLO E PLANTA
pendentes_solo_df<- filter(pendentes_df, Laboratório == "SOLO E PLANTA")%>%
  select(Lote, Cliente, Fazenda, `Nº Amostras`, 
         `Data Envio`, `Data Recebimento`,
         `Previsão de Liberação`)  

pendentes_ibra_df<- filter(pendentes_df, Laboratório == "IBRA")%>%
  select(Lote, Cliente, Fazenda, `Nº Amostras`, 
         `Data Envio`, `Data Recebimento`,
         `Previsão de Liberação`)  

pendentes_solo_total <- count(pendentes_solo_df$Fazenda)
pendentes_solo_total2<-sum(pendentes_solo_total$freq)
total_solo_lote<-paste0( "Nº de lotes pendentes: ", pendentes_solo_total2)

pendentes_solo_total3<-sum(pendentes_solo_df$`Nº Amostras`)
total_solo_amostra<-paste0( "Nº de amostras pendentes: ", pendentes_solo_total3)


#JP EM AGUARDO
jp_df <- read_excel("G:/.shortcut-targets-by-id/1eY6qLaqCw0vMplswsEaSexf_xogKRn_4/GEOPROCESSAMENTO/RESUMO DIARIO - 2022/Controle Geral Amostras - 2022_v2_geo.xlsm", sheet = "JP OUTROS")


###############################################################################
#                   EM AGUARDO PARA PROCESSAMENTO
###############################################################################

aguardo_processamento <- filter(df, `Status Processamento` =="EM AGUARDO")%>%
  select(Cliente, Fazenda, `Nº Amostras`, `Data de Liberação`)

aguardo_processamento$`Data de Liberação` <- as.Date(aguardo_processamento$`Data de Liberação`)

aguardo_df1<-ddply(aguardo_processamento, .(Cliente, Fazenda), summarize,
                   `Nº amostras aguardo` = round(sum(`Nº Amostras`), 2))

Pendentes_df1<-ddply(pendentes_df, .(Cliente, Fazenda), summarize,
                     `Nº amostras Pendentes` = round(sum(`Nº Amostras`), 2))

data<-ddply(aguardo_processamento, .(Cliente, Fazenda), summarize, `Data ultimo resultado` =  max(`Data de Liberação`))


join1<-join(aguardo_df1, Pendentes_df1,  type = "left" )# Primeiro Join Cliente + fazenda + amostras pendentes
join1[is.na(join1)] <- 0
join2<-join(join1, data, type = "left") # segundo join cliente + fazenda + amostras pendentes + data

options(digits = 3)

join2["Nº total amostras"]<-join2$`Nº amostras aguardo` + join2$`Nº amostras Pendentes` # total de amostra da fazenda
join2["Status de conclusão (%)"]<-(join2$`Nº amostras aguardo` * 100) / join2$`Nº total amostras`  # status de conclusão


aguardo_total <- count(join2$Fazenda)
aguardo_total2<-sum(aguardo_total$freq)

total_aguardo<-paste0( "Nº Em Aguardo: ", aguardo_total2)


###############################################################################
#                          PROCESSADOS DO DIA
###############################################################################

processados_df <- read_excel("G:/.shortcut-targets-by-id/1eY6qLaqCw0vMplswsEaSexf_xogKRn_4/GEOPROCESSAMENTO/RESUMO DIARIO - 2022/Controle Geral Amostras - 2022_v2_geo.xlsm", sheet = "PROCESSADOS")

processados_total <- count(processados_df$Fazenda)
t <- sum(processados_total$freq)

aguardo_total2<-sum(aguardo_total$freq)
total_processado<-paste0( "Nº de Processamentos do dia: ", t)



###############################################################################
#                           CRIANDO RESUMO DOC
###############################################################################


# LENDO MODELO DOCX
sample_doc<-read_docx("G:/.shortcut-targets-by-id/1eY6qLaqCw0vMplswsEaSexf_xogKRn_4/GEOPROCESSAMENTO/RESUMO DIARIO - 2022/MODELO.docx")
sample_doc<-  cursor_end(sample_doc) %>% body_remove()
sample_doc<- body_add_fpar(sample_doc,fpar(ftext(DATA_ATUAL, prop = text_style_2)))


# IBRA PREVISAO

IBRA <- flextable(pendentes_ibra_df) 



footer_str <- paste0(total_ibra_lote, sep="\n", total_ibra_amostra)
header_str <- 'IBRA - AMOSTRAS DE SOLO'

IBRA <- add_header_lines(IBRA, values = header_str)
IBRA <- add_footer_lines(IBRA, values = footer_str)
IBRA <- theme_zebra(IBRA)
IBRA <- bg(IBRA, i= 1, part = "header", bg = "#2E8B57")
IBRA <- bg(IBRA, i= 2, part = "header", bg = "#3CB371")
IBRA <- bg(IBRA, i= 1, part = "footer", bg = "#2E8B57")
IBRA <- align(IBRA, align = "center", part="all")
IBRA <- flextable::width(IBRA,width=1.5)
set_table_properties(IBRA, layout = "autofit")


IBRA<-fontsize(IBRA, size = 10, part = "body")


sample_doc <- flextable::body_add_flextable(sample_doc, value = IBRA, keepnext = FALSE)

sample_doc <- officer::body_end_block_section(sample_doc, value=landscape )

# SOLO E PLANTA

SOLO <- flextable(pendentes_solo_df)

footer_str <- paste0(total_solo_lote, sep="\n", total_solo_amostra)
header_str <- 'SOLO E PLANTA - AMOSTRAS DE SOLO'

SOLO <- add_header_lines(SOLO, values = header_str)
SOLO <- add_footer_lines(SOLO, values = footer_str)
SOLO <- theme_zebra(SOLO)
SOLO <- bg(SOLO, i= 1, part = "header", bg = "#FFD700")
SOLO <- bg(SOLO, i= 2, part = "header", bg = "#F0E68C")
SOLO <- bg(SOLO, i= 1, part = "footer", bg = "#F0E68C")
SOLO <- align(SOLO, align = "center", part="all")
SOLO <- flextable::width(SOLO,width=1.5)
set_table_properties(SOLO, layout = "autofit")
SOLO<-fontsize(SOLO, size = 10, part = "body")

sample_doc <- flextable::body_add_flextable(sample_doc, value = SOLO)
sample_doc <- officer::body_end_block_section(sample_doc, value=landscape)



#                             EXATA PREVISAO

EXATA <- flextable(pendentes_exata_df)
#sample_doc <- officer::body_end_block_section(sample_doc, value=landscape)
#sample_doc<-body_add_par(sample_doc,"", pos = "before")

footer_str <- paste0(total_exata_lote, sep="\n", total_exata_amostra)
header_str <- 'EXATA - AMOSTRAS DE SOLO'
EXATA <- add_header_lines(EXATA, values = header_str)
EXATA <- add_footer_lines(EXATA, values = footer_str)
EXATA <- theme_zebra(EXATA)
EXATA <- align(EXATA, align = "center", part="all")
EXATA <- bg(EXATA, i= 1, part = "header", bg = "#6495ED")
EXATA <- bg(EXATA, i= 2, part = "header", bg = "#ADD8E6")
EXATA <- bg(EXATA, i= 1, part = "footer", bg = "#ADD8E6")
EXATA <- flextable::width(EXATA,width=1.3)
set_table_properties(EXATA, layout = "autofit")
EXATA<-fontsize(EXATA, size = 10, part = "body")
sample_doc <- flextable::body_add_flextable(sample_doc, value = EXATA) 
#doc <- body_add_break(sample_doc)


# Em aguardo para processamento
total_processado<-paste0( "Nº de Processamentos do dia: ", t)

join <- flextable(join2)
sample_doc <- officer::body_end_block_section(sample_doc, value=landscape)

#footer_str <- 'AMOSTRAS DE SOLO'
header_str <- 'EM AGUARDO PARA PROCESSAMENTO'
footer_str <- paste0('Nº Amostras Pendentes : refere-se ao número de amostras ainda sem resultados', sep="\n", 
                     'Status de conclusão (%): refere-se a porcentagem de conclusão de resultados liberados pelos laboratórios', 
                     sep="\n", total_aguardo)
join <- add_header_lines(join, values = header_str)
join <- add_footer_lines(join, values = footer_str)
join <- theme_zebra(join)
join <- bg(join, i= 1, part = "header", bg = "#FFD700")
join <- bg(join, i= 2, part = "header", bg = "#F0E68C")
join <- bg(join, i= 1, part = "footer", bg = "#FFD700")
join <- align(join, align = "center", part="all")
join <- flextable::width(join,width=1.5)
set_table_properties(join, layout = "autofit")
join<-fontsize(join, size = 10, part = "body") 
#colourer <- col_numeric(palette = "RdYlGn",domain = c(0, 100))
#join<-bg(join, j = "Status de conclusão (%)", bg = colourer)

sample_doc <- flextable::body_add_flextable(sample_doc, value = join)
sample_doc <- officer::body_end_block_section(sample_doc, value=landscape)
#doc <- body_add_break(sample_doc)



# JP EM AGUARDO

JP <- flextable(jp_df)

#sample_doc <- officer::body_end_block_section(sample_doc, value=landscape)
sample_doc<-body_add_par(sample_doc,"", pos = "after")


#footer_str <- 'AMOSTRAS DE SOLO'
header_str <- 'JP e Outros - Em aguardo para processamento'
JP <- add_header_lines(JP, values = header_str)
#table <- add_footer_lines(table, values = footer_str)
JP <- theme_zebra(JP)
JP <- bg(JP, i= 1, part = "header", bg = "#7B68EE")
JP <- bg(JP, i= 2, part = "header", bg = "#9370DB")
JP <- align(JP, align = "center", part="all")
JP <- flextable::width(JP,width=2.8)
set_table_properties(JP, layout = "autofit")
JP<-fontsize(JP, size = 10, part = "body")

sample_doc <- flextable::body_add_flextable(sample_doc, value = JP)

sample_doc <- officer::body_end_block_section(sample_doc, value=portrait)

# PROCESSADOS DO DIA

processados_df <- flextable(processados_df)

footer_str <- total_processado
header_str <- 'PROCESSADOS'
processados_df <- add_header_lines(processados_df, values = header_str)
processados_df <- add_footer_lines(processados_df, values = footer_str)
processados_df <- theme_zebra(processados_df)
processados_df <- bg(processados_df, i= 1, part = "header", bg = "#4F4F4F")
processados_df <- bg(processados_df, i= 1, part = "footer", bg = "#4F4F4F")
processados_df <- align(processados_df, align = "center", part="all")
processados_df <- flextable::width(processados_df,width=2.8)
set_table_properties(processados_df, layout = "autofit")
processados_df<-fontsize(processados_df, size = 10, part = "body")

#sample_doc <- officer::body_end_block_section(sample_doc, value=portrait)
sample_doc <- flextable::body_add_flextable(sample_doc, value = processados_df)


path<-"G:/.shortcut-targets-by-id/1eY6qLaqCw0vMplswsEaSexf_xogKRn_4/GEOPROCESSAMENTO/RESUMO DIARIO - 2022/"

print(sample_doc, target = paste0(path,DATA_ATUAL,".docx") )


