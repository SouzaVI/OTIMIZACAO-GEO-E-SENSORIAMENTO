#-------------------------------------------------------------------------------
# LOOP INTERPOLACAO
# Autor: Igor Viana Souza ( M.e Engenharia Florestal)
# Data de criaCAo: 01/10/2021
#-------------------------------------------------------------------------------

#                         LIMPAR MEMORIA
#-------------------------------------------------------------------------------

rm(list=ls())

#-------------------------------------------------------------------------------
#                           PACOTES
#-------------------------------------------------------------------------------

if(!require(geobr)){install.packages("geobr")}
if(!require(readxl)){install.packages("readxl")}
if(!require(dplyr)){install.packages("dplyr")}
if(!require(sf)){install.packages("sf")}
if(!require(gstat)){install.packages("gstat")}
if(!require(ggplot2)){install.packages("ggplot2")}
if(!require(fields)){install.packages("fields")}
if(!require(magick)){install.packages("magick")}
if(!require(purrr)){install.packages("purrr")}


#-------------------------------------------------------------------------------
#                           DADOS
#-------------------------------------------------------------------------------

temp <- read_excel("3_INTERPOLACAO.xlsx")

lista_temp<-list()
for (a in temp$Data) {
  lista_temp[[a]]<-filter(temp, Data == a) %>% st_as_sf(coords = c('long', 'lat'), 
                                                        crs=4674)
}
#-------------------------------------------------------------------------------
#                           GRADE
#-------------------------------------------------------------------------------

go<-read_state(code_state = "TO")
muni<-read_municipality(
  code_muni = "TO",
  year = 2010,
  simplified = TRUE,
  showProgress = TRUE
)
grade_go<-st_make_grid(go, cellsize = c(.3,.3)) %>% #TAMANHO DA GRADE
  st_as_sf()%>%
  filter(st_contains(go,., sparse = FALSE))
#-------------------------------------------------------------------------------
#                         GERANDO MODELO
#-------------------------------------------------------------------------------
modelo_list<-list()

for (x in 1:length(lista_temp)) {
  modelo_list[[x]]<-gstat(formula = Umidade~1,
                          data=as(lista_temp[[x]],'Spatial'),
                          set=list(idp=3))
}
#-------------------------------------------------------------------------------
#                            PREVISAO
#-------------------------------------------------------------------------------
temp.int<-list()

for (y in 1:length(modelo_list)) {
  temp.int[[y]]<-predict(modelo_list[[y]],as(grade_go,'Spatial'))%>%st_as_sf()


}

#

date_name<-names(lista_temp)
names(temp.int)<- c(date_name)


#-------------------------------------------------------------------------------  
#                         GERANDO MAPAS
#-------------------------------------------------------------------------------
for (i in 1:length(temp.int)) {
  
names(temp.int)<- c(date_name)

map<-list()
map[[i]]<-ggplot(temp.int[[i]])+
    geom_sf(aes(fill=var1.pred,col=var1.pred))+
    geom_sf(data=go, fill= 'transparent')+
    geom_sf(data=muni, fill= 'transparent', colour = "black")+
    scale_color_gradientn(colors = tim.colors(50), 
                          limits=c(0,80))+
    scale_fill_gradientn(colors = tim.colors(50), 
                         limits=c(0,80))+
    theme_void()+
    labs(
         fill="%",
         color="%",
         subtitle = paste0(i,"/202X"))+
  ggtitle(paste0("Úmidade Relativa (%)"))

ggsave(map[[i]], file=paste0("plot_", i,".png"), width = 14, height = 10, units = "cm")
}


#-------------------------------------------------------------------------------  
#                       EXPORTACAO 
#-------------------------------------------------------------------------------

iwalk(temp.int, function(dat, name){
  st_write(obj = dat, dsn = paste0("20XX_", name, ".shp"))
})
