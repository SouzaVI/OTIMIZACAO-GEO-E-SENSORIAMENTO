#-------------------------------------------------------------------------------
# Gerar apresentacao automatica, inserir mapas e média com base no dbf
# Autor: Igor Viana Souza ( M.e Engenharia Florestal)
# Data de criacaoo: 01/05/2022
#-------------------------------------------------------------------------------


rm(list=ls(all=TRUE))

if(!require(officer)){install.packages("officer")}
if(!require(rgdal)){install.packages("rgdal")}
if(!require(tidyverse)){install.packages("tidyverse")}
if(!require(dplyr)){install.packages("dplyr")}
if(!require(sf)){install.packages("sf")}
if(!require(foreign)){install.packages("foreign")}



MASTER<-read_pptx("C:/@IGOR/R/MAPAS/modelo.pptx")
layout<-on_slide(MASTER, index= 2)
#0-10
argila2<-on_slide(MASTER, index= 3)
mo2<-on_slide(MASTER, index= 4)
ctc2<-on_slide(MASTER, index= 5)
ph2<-on_slide(MASTER, index= 6)
v2<-on_slide(MASTER, index= 7)
ca2<-on_slide(MASTER, index= 8)
mg2<-on_slide(MASTER, index= 9)
k2<-on_slide(MASTER, index= 10)
p2<-on_slide(MASTER, index= 11)
al2<-on_slide(MASTER, index= 12)
s2<-on_slide(MASTER, index= 13)
b2<-on_slide(MASTER, index= 14)
cu2<-on_slide(MASTER, index= 15)
mn2<-on_slide(MASTER, index= 16)
zn2<-on_slide(MASTER, index= 17)
#10-20
mo3<-on_slide(MASTER, index= 18)
ctc3<-on_slide(MASTER, index= 19)
ph3<-on_slide(MASTER, index= 20)
v3<-on_slide(MASTER, index= 21)
ca3<-on_slide(MASTER, index= 22)
mg3<-on_slide(MASTER, index= 23)
k3<-on_slide(MASTER, index= 24)
p3<-on_slide(MASTER, index= 25)
al3<-on_slide(MASTER, index= 26)
#20-40
ctc4<-on_slide(MASTER, index= 32)
v4<-on_slide(MASTER, index= 33)
ca4<-on_slide(MASTER, index= 34)
al4<-on_slide(MASTER, index= 35)
s4<-on_slide(MASTER, index= 36)

##ATRIBUTOS FERTILIDADE SOLO

FERTILIDADE_DTF<-read.dbf("C:/@IGOR/PYTHON/MAPA/app.dbf", as.is = FALSE)


##MAPAS FERTILIDADE

MASTER_2 <- ph_with(layout, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/LAYOUT._FERTILIDADE_LAYOUT.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(argila2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/ARGILA2._FERTILIDADE_ARGILA2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(mo2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/MO2._FERTILIDADE_MO2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ctc2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CTC2._FERTILIDADE_CTC2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ph2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/PH2._FERTILIDADE_PH2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(v2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/V2._FERTILIDADE_V2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ca2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CA2._FERTILIDADE_CA2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(mg2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/MG2._FERTILIDADE_MG2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(k2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/K2._FERTILIDADE_K2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(p2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/P2._FERTILIDADE_P2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(al2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/AL2._FERTILIDADE_AL2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(s2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/S2._FERTILIDADE_S2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(b2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/B2._FERTILIDADE_B2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(cu2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CU2._FERTILIDADE_CU2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(mn2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/MN2._FERTILIDADE_MN2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(zn2, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/ZN2._FERTILIDADE_ZN2.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )


##10-20 CM

MASTER_2 <- ph_with(argila3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/ARGILA3._FERTILIDADE_ARGILA3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(mo3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/MO3._FERTILIDADE_MO3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ctc3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CTC3._FERTILIDADE_CTC3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ph3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/PH3._FERTILIDADE_PH3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(v3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/V3._FERTILIDADE_V3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(ca3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CA3._FERTILIDADE_CA3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(mg3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/MG3._FERTILIDADE_MG3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(k3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/K3._FERTILIDADE_K3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(p3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/P3._FERTILIDADE_P3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )
MASTER_2 <- ph_with(al3, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/AL3._FERTILIDADE_AL3.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )

##20-40 CM

MASTER_2 <- ph_with(ctc4, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CTC4._FERTILIDADE_CTC4.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )

MASTER_2 <- ph_with(v4, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/V4._FERTILIDADE_V4.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )

MASTER_2 <- ph_with(ca4, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/CA4._FERTILIDADE_CA4.png", width = 8.0, height = 6.0), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )

MASTER_2 <- ph_with(al4, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/AL4._FERTILIDADE_AL4.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )

MASTER_2 <- ph_with(s4, external_img("C:/@IGOR/PYTHON/MAPA/PROJETOS/S4._FERTILIDADE_S4.png", width = 8.0, height = 6), res = 800,
                    units = "cm", alignment = "c",
                    location = ph_location_type(type = "body"), use_loc_size = FALSE )


# 0-10 medias
argila2_df<-mean(FERTILIDADE_DTF$ARGILA2)%>%round(2)
mo2_df<-mean(FERTILIDADE_DTF$MO2)%>%round(2)
ctc2_df<-mean(FERTILIDADE_DTF$CTC2)%>%round(2)
ph2_df<-mean(FERTILIDADE_DTF$PH2)%>%round(2)
v2_df<-mean(FERTILIDADE_DTF$V2)%>%round(2)
ca2_df<-mean(FERTILIDADE_DTF$CA2)%>%round(2)
mg2_df<-mean(FERTILIDADE_DTF$MG2)%>%round(2)
k2_df<-mean(FERTILIDADE_DTF$K2)%>%round(2)
p2_df<-mean(FERTILIDADE_DTF$P2)%>%round(2)
al2_df<-mean(FERTILIDADE_DTF$AL2)%>%round(2)
s2_df<-mean(FERTILIDADE_DTF$S2)%>%round(2)
b2_df<-mean(FERTILIDADE_DTF$B2)%>%round(2)
cu2_df<-mean(FERTILIDADE_DTF$CU2)%>%round(2)
mn2_df<-mean(FERTILIDADE_DTF$MN2)%>%round(2)
zn2_df<-mean(FERTILIDADE_DTF$ZN2)%>%round(2)
#10-20 media
mo3_df<-mean(FERTILIDADE_DTF$MO3)%>%round(2)
ctc3_df<-mean(FERTILIDADE_DTF$CTC3)%>%round(2)
ph3_df<-mean(FERTILIDADE_DTF$PH3)%>%round(2)
v3_df<-mean(FERTILIDADE_DTF$V3)%>%round(2)
ca3_df<-mean(FERTILIDADE_DTF$CA3)%>%round(2)
mg3_df<-mean(FERTILIDADE_DTF$MG3)%>%round(2)
k3_df<-mean(FERTILIDADE_DTF$K3)%>%round(2)
p3_df<-mean(FERTILIDADE_DTF$P3)%>%round(2)
al3_df<-mean(FERTILIDADE_DTF$AL3)%>%round(2)
#20-40 media
ctc4_df<-mean(FERTILIDADE_DTF$CTC4)%>%round(2)
v4_df<-mean(FERTILIDADE_DTF$V4)%>%round(2)
ca4_df<-mean(FERTILIDADE_DTF$CA4)%>%round(2)
al4_df<-mean(FERTILIDADE_DTF$AL4)%>%round(2)
s4_df<-mean(FERTILIDADE_DTF$S4)%>%round(2)




# 0-10 slide
MASTER_2 <- ph_with(argila2, value = argila2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(mo2, value = mo2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ctc2, value = ctc2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ph2, value = ph2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(v2, value = v2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ca2, value = ca2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(mg2, value = mg2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(k2, value = k2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(p2, value = p2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(al2, value = al2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(s2, value = s2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(b2, value = b2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(cu2, value = cu2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(mn2, value = mn2_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(zn2, value = zn2_df, location = ph_location_type(type = "dt"))



#10-20 slide
MASTER_2 <- ph_with(argila3, value = argila3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(mo3, value = mo3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ctc3, value = ctc3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ph3, value = ph3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(v3, value = v3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ca3, value = ca3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(mg3, value = mg3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(k3, value = k3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(p3, value = p3_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(al3, value = al3_df, location = ph_location_type(type = "dt"))
#20-40 slide
MASTER_2 <- ph_with(ctc4, value = ctc4_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(v4, value = v4_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(ca4, value = ca4_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(al4, value = al4_df, location = ph_location_type(type = "dt"))
MASTER_2 <- ph_with(s4, value = s4_df, location = ph_location_type(type = "dt"))



#organizador slide

MASTER_2<-move_slide(MASTER_2, index = 18, to = 5)
MASTER_2<-move_slide(MASTER_2, index = 19, to = 7)
MASTER_2<-move_slide(MASTER_2, index = 32, to = 8)
MASTER_2<-move_slide(MASTER_2, index = 21, to = 10)
MASTER_2<-move_slide(MASTER_2, index = 22, to = 12)
MASTER_2<-move_slide(MASTER_2, index = 33, to = 13)
MASTER_2<-move_slide(MASTER_2, index = 24, to = 15)
MASTER_2<-move_slide(MASTER_2, index = 34, to = 16)
MASTER_2<-move_slide(MASTER_2, index = 26, to = 18)
MASTER_2<-move_slide(MASTER_2, index = 27, to = 20)
MASTER_2<-move_slide(MASTER_2, index = 28, to = 22)
MASTER_2<-move_slide(MASTER_2, index = 29, to = 24)
MASTER_2<-move_slide(MASTER_2, index = 35, to = 25)
MASTER_2<-move_slide(MASTER_2, index = 36, to = 27)




#exportar
print(MASTER_2, target = "C:/@IGOR/R/saida.pptx")


