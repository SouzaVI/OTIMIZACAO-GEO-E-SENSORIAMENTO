#-------------------------------------------------------------------------------
# Organizador de resultados EXATA
# Autor: Igor Viana Souza ( M.e Engenharia Florestal)
# Data de cria??o: 01/10/2021
#-------------------------------------------------------------------------------
#                         LIMPAR MEMORIA
#-------------------------------------------------------------------------------

rm(list=ls(all=TRUE))

#-------------------------------------------------------------------------------
#                           PACOTES
#-------------------------------------------------------------------------------

if(!require(dplyr)){install.packages("dplyr")}
if(!require(readxl)){install.packages("readxl")}
if(!require(openxlsx)){install.packages("openxlsx")}

#-------------------------------------------------------------------------------
#                       CARREGANDO DADOS
#-------------------------------------------------------------------------------

EXATA <- read_excel("C:/@IGOR/R/exata.xls",  skip = 6)## SERVIR? PARA EXTRAIR O ID
EXATA2 <- EXATA ## SERVIR? PARA EXTRAIR PROF
EXATA["Prof"] <-c()# ADD COLUNA DE Prof

#-------------------------------------------------------------------------------
#                   FUNCAO SEPARAR ID E PROFUNDIDADE
#-------------------------------------------------------------------------------

separar <- function(x,qual=1,sep=" ") {
  #pega o texto e quebra segundo o separador informado
  xx = strsplit(x,split=sep)[[1]]
  #retorna o valor
  return(xx[qual])}
  
#-------------------------------------------------------------------------------
#                                 PEGAR ID

lapply(EXATA$id,separar,qual=1)
id = as.vector(lapply(EXATA$id,separar,qual=1),mode='double')
id 
#ADICIONAMOS NO OBJETO ORIGINAL

EXATA$id = id

#                             PEGAR PROFUNDIDADE

prof = as.vector(lapply(EXATA2$id,separar,qual=2),mode='character')

#                      ADICIONAMOS NO OBJETO ORIGINAL

EXATA$Prof = prof

#                    ORGANIZANDO CABE?ALHO E REMOVENDO NS

a<-EXATA %>% dplyr::select(Lab, id, Prof, Zn, Mn, Fe, Cu, B, S,`Sat. Al`, Al, 
                        `P(meh)`,`P (rem)`, `P (res)`,
                        `K/CTC`, K1, `Ca/Mg`, `Ca/Mg`, `Mg/CTC`, Mg, `Ca/CTC`,
                        Ca, `Sat. Bases`,`ph CaCl2`, CTC, `Mat. Org.`, Argila )
   names(a)[1:26] <- c("Lab","id","prof","zn2",	"mn2",	"fe2",	"cu2",	"b2",	
                       "s2",	"sat_al2",	"al2",	"p_mel2",	"p_rem2",	"p_res2",	
                       "sat_k2",	"k2",	"rel_ca_mg2",	"sat_mg2",	"mg2",	"sat_ca2",
                       "ca2",	"v2",	"ph2",	"ctc2",	"mo2",	"argila2")

a[a == "ns"] <-NA 

#                               DATAFRAME PARA N?MEROS
b<-a %>% 
  mutate_at(c("zn2",	"mn2",	"fe2",	"cu2",	"b2",	"s2",	"sat_al2",	"al2",	
              "p_mel2",	"p_rem2",	"p_res2",	
              "sat_k2",	"k2",	"rel_ca_mg2",	"sat_mg2",	"mg2",	"sat_ca2",
              "ca2",	"v2",	"ph2",	"ctc2",	"mo2",	"argila2"), 
            funs(as.numeric(as.character(.))))
round(b[4:26],2)

##                               FILTRO POR PROFUNDIDADE
pro_010 <- b%>%filter(prof=="(0-10)")
names(pro_010)[1:26] <- c("Lab","id","prof","zn2",	"mn2",	"fe2",	"cu2",	"b2",	
                          "s2",	"sat_al2",	"al2",	"p_mel2",	"p_rem2",	"p_res2",	
                    "sat_k2",	"k2",	"rel_ca_mg2",	"sat_mg2",	"mg2",	"sat_ca2",
                    "ca2",	"v2",	"ph2",	"ctc2",	"mo2",	"argila2")
###
pro_1020 <- b%>%filter(prof=="(10-20)")
names(pro_1020)[1:26] <- c("Lab","id","prof","zn3",	"mn3",	"fe3",	"cu3",	"b3",
                           "s3",	"sat_al3",	"al3",	"p_mel3",	"p_rem3",	"p_res3",	
                          "sat_k3",	"k3",	"rel_ca_mg3",	"sat_mg3",	"mg3",	"sat_ca3",
                          "ca3",	"v3",	"ph3",	"ctc3",	"mo3",	"argila3")
###
pro_2040 <- b%>%filter(prof=="(20-40)")
names(pro_2040)[1:26] <- c("Lab","id","prof","zn4",	"mn4",	"fe4",	"cu4",	"b4",
                           "s4",	"sat_al4",	"al4",	"p_mel4",	"p_rem4",	"p_res4",	
                           "sat_k4",	"k4",	"rel_ca_mg4",	"sat_mg4",	"mg4",	
                           "sat_ca4",
                           "ca4",	"v4",	"ph4",	"ctc4",	"mo4",	"argila4")


##
pro_020 <- b%>%filter(prof=="(0-20)")
names(pro_020)[1:26] <- c("Lab","id","prof","zn1",	"mn1",	"fe1",	"cu1",	"b1",	
                          "s1",	"sat_al1",	"al1",	"p_mel1",	"p_rem1",	"p_res1",	
                           "sat_k1",	"k1",	"rel_ca_mg1",	"sat_mg1",	"mg1",	
                          "sat_ca1",
                           "ca1",	"v1",	"ph1",	"ctc1",	"mo1",	"argila1")

##                                   EXPORTANDO XLSX
list_of_datasets <- list("Geral" = b, "0-10" = pro_010, "10-20"= pro_1020, 
                         "20-40" = pro_2040, "0-20" = pro_020 )
write.xlsx(list_of_datasets, "C:/@IGOR/R/Exata1.xlsx", overwrite = TRUE)

