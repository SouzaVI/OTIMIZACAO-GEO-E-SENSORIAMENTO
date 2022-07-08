#-------------------------------------------------------------------------------
# CALCULO NDVI, MASK RGB E NDVI
# Autor: Igor Viana Souza ( M.e Engenharia Florestal)
# Data de cria??o: 01/10/2021
#-------------------------------------------------------------------------------

#                         LIMPAR MEMORIA
#-------------------------------------------------------------------------------

rm(list=ls(all=TRUE))

#-------------------------------------------------------------------------------
#                           PACOTES
#-------------------------------------------------------------------------------

#install.packages('terra', repos='https://rspatial.r-universe.dev')
if(!require(raster)){install.packages("raster")}
if(!require(rgdal)){install.packages("rgdal")}
#-------------------------------------------------------------------------------
#                  DIRETORIO E ARMAZENAMENTO
#-------------------------------------------------------------------------------

setwd("C:/@IGOR/R/")                                                  
RGB <- list.files(path="RGB", pattern = ".tif", full.names = TRUE)    
SHP <- list.files(path="TALHAO", pattern = ".shp", full.names = TRUE)
INDICE <- list.files(path="NDVI", pattern = ".tif", full.names = TRUE) 

#-------------------------------------------------------------------------------
#                        LENDO RGB E NDVI
#-------------------------------------------------------------------------------
RGB_STACK<-c()                                                        
r_name <- list.files(path="RGB",full.names = F)

for(i in RGB){ RGB_STACK[[i]]<- stack(i)} 

names(RGB_STACK)<- c(r_name)
#-------------------------------------------------------------------------------
#                              NDVI
#-------------------------------------------------------------------------------
NDVI_STACK<-c() 

r_name1 <- list.files(path="NDVI",full.names = F)

for(g in INDICE){ 
  NDVI_STACK[[g]]<- stack(g)} 


names(NDVI_STACK)<- c(r_name1)

bandas<-NDVI_STACK[[1]] ##ALTERE O N?MERO, PARA MUDAR A IMAGEM, LEMBRANDO QUE ESSE N?MERO
## DEVERA SE REPETIR NAS LINHAS EM QUE FOR SOLICITADO
## TRATA-SE DE UM FORMA DE SELECIONAR AS IMAGENS NA PASTA

#-------------------------------------------------------------------------------
#                           CALCULO NDVI
#-------------------------------------------------------------------------------

sat_ndvi_list<-list()

sat_ndvi <- (bandas[[1]] - bandas[[3]]) / (bandas[[1]] + bandas[[3]])
plot(sat_ndvi)

NDVI_STACK<- stack(sat_ndvi) 

#-------------------------------------------------------------------------------                                                                  
#                           LENDO TALHAO
#-------------------------------------------------------------------------------

TALHAO<-c()                                                           
for(j in SHP){TALHAO[[j]]<- sf::st_read(j)}                           
  
#-------------------------------------------------------------------------------                                                                    
#                           LOOP MASK RGB
#-------------------------------------------------------------------------------

CO<-TALHAO[[1]]

TA<-CO$TALHAO

output <- data.frame(t = TA)


dataoutput<-data.frame(data=r_name[1]) ##ALTERA O N?MERO, PARA MUDAR A IMAGEM

names_data<-paste("RGB_",output$t,dataoutput$data)

LISTA<-list()


RGB_STACK
for (c in 1:length(output$t)) {LISTA[c]<- mask(RGB_STACK[[1]], CO[CO$TALHAO==output$t[c],])
  
}## ALTERE O N?MERO EM RGB[[X]], PARA MUDAR A IMAGEM, ESSE NUMERO DEVE SER O MESMO ANTERIOR
names(LISTA)<- c(names_data)

#names(LISTA)<- c(CO$TALHAO)

#-------------------------------------------------------------------------------
#                              LOOP MASK NDVI
#-------------------------------------------------------------------------------

LISTA_NDVI<-list()

for (y in 1:length(output$t)) {LISTA_NDVI[y]<- mask(sat_ndvi, CO[CO$TALHAO==output$t[y],])
}

names_data2<-paste("NDVI_",output$t,dataoutput$data)
names(LISTA_NDVI)<- c(names_data2)
#names(LISTA_NDVI)<- c(CO1$TALHAO)

#-------------------------------------------------------------------------------
mapply(writeRaster, LISTA, names(LISTA), 'GTiff',overwrite= TRUE)
mapply(writeRaster, LISTA_NDVI, names(LISTA_NDVI), 'GTiff', overwrite= TRUE)





