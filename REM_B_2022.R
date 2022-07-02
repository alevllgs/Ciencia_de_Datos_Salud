# Librerias ---------------------------------------------------------------
library(tidyverse)
library(readxl)
library(lubridate)
library(janitor)
library(dplyr)
library(openxlsx)
library(xlsx)

# Archivos mensuales ------------------------------------------------------
#__________Cada mes debo cambiar las variables #fecha_mes y archivo

meses <- c("02")

# meses <- c("06", "07", "08", "09", "10", "11","12")
for (i in meses) {

fecha_mes <- paste0("2022-",i,"-01")  
archivoBS <- paste0("REM/2022-",i," REM serie BS.xlsx")

# Representan las BBDD donde esta guardada la información -----------------

B_IMGBBDD <- "BBDD/B_IMG BBDD.xlsx"
B_LABBBDD <- "BBDD/B_LAB BBDD.xlsx"
B_APBBDD <- "BBDD/B_AP BBDD.xlsx"
B_UMTBBDD <- "BBDD/B_UMT BBDD.xlsx"
B_QfBBDD <- "BBDD/B_Qf BBDD.xlsx"
B171_QfBBDD <- "BBDD/B171_Qf BBDD.xlsx"
B172_QfBBDD <- "BBDD/B172_Qf BBDD.xlsx"

# Imagenología ------------------------------------------------------

B_Img01 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A750:AL806")
B_Img01$'Clasificación' <- "Ex Radiologicos Simples"
B_Img02 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A809:AL829")
B_Img02$'Clasificación' <- "Ex Radiologicos Complejos"
B_Img03 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A832:AL859")
B_Img03$'Clasificación' <- "Tomografia Axial Comp"
B_Img04 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A862:AL884")
B_Img04$'Clasificación' <- "Ultrasonografia"
B_Img05 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A887:AL925")
B_Img05$'Clasificación' <- "Resonancia Magnética"
B_ImgM <- rbind(B_Img01, B_Img02, B_Img03, B_Img04, B_Img05)
B_ImgM <- mutate_all(B_ImgM, ~replace(., is.na(.), 0))
B_ImgM$Fecha <- fecha_mes
B_ImgM$"Centros de Costos" <-'542-IMAGENOLOGÍA'
colnames(B_ImgM)[1] <- "Codigo"
colnames(B_ImgM)[2] <- "Glosa"
colnames(B_ImgM)[3] <- "Total"
colnames(B_ImgM)[27] <- "Procedencia Atención Cerrada"
colnames(B_ImgM)[28] <- "Procedencia Atención Abierta"
colnames(B_ImgM)[29] <- "Procedencia Emergencia"
colnames(B_ImgM)[38] <- "Total Facturado"
B_ImgM <- B_ImgM %>%select(Fecha,"Centros de Costos", 'Clasificación', Codigo,Glosa, Total, `Total Facturado`, `Procedencia Atención Abierta`, `Procedencia Atención Cerrada`, `Procedencia Emergencia`)
rm(B_Img01,B_Img02,B_Img03,B_Img04,B_Img05) #borro los objetos innecesarios
B_ImgM$"AT Cerrada" <- ifelse(B_ImgM$`Total Facturado`==0,0,
                             B_ImgM$`Total Facturado`*(B_ImgM$`Procedencia Atención Cerrada`/B_ImgM$Total))

# Laboratorio -------------------------------------------------------------
B_Lab01 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A224:AL296")
B_Lab01$'Clasificación' <- "I Sangre, Hematología"
B_Lab02 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A299:AL374")
B_Lab02$'Clasificación' <- "II Sangre, Ex Bioquimicos"
B_Lab03 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A377:AL420")
B_Lab03$'Clasificación' <- "III Hormonas"
B_Lab04 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A423:AL434")
B_Lab04$'Clasificación' <- "IV Genetica"
B_Lab05 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A437:AL513")
B_Lab05$'Clasificación' <- "V Inmunologia"
B_Lab06 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A517:AL572")
B_Lab06$'Clasificación' <- "VI Ex Microbiologicos (bacterias y hongos)"
B_Lab07 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A575:AL592")
B_Lab07$'Clasificación' <- "VI Ex Microbiologicos (parasitos)"
B_Lab08 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A595:AL633")
B_Lab08$'Clasificación' <- "VI Ex Microbiologicos (virus)"
B_Lab09 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A636:AL643")
B_Lab09$'Clasificación' <- "VII Proc o determinaciones directamente con el paciente"
B_Lab10 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A657:AL704")
B_Lab10$'Clasificación' <- "VIII Ex de deposiciones, exudados, secreciones y otros liquidos"
B_Lab11 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A707:AL746")
B_Lab11$'Clasificación' <- "IX Examenes de orina"

B_LabM <- rbind(B_Lab01, B_Lab02, B_Lab03, B_Lab04, B_Lab05,B_Lab06, B_Lab07, B_Lab08, B_Lab09, B_Lab10,B_Lab11)
B_LabM <- mutate_all(B_LabM, ~replace(., is.na(.), 0))
B_LabM$Fecha <- fecha_mes
B_LabM$"Centros de Costos" <-'518-LABORATORIO CLÍNICO'

colnames(B_LabM)[1] <- "Codigo"
colnames(B_LabM)[2] <- "Glosa"
colnames(B_LabM)[3] <- "Total"
colnames(B_LabM)[27] <- "Procedencia Atención Cerrada"
colnames(B_LabM)[28] <- "Procedencia Atención Abierta"
colnames(B_LabM)[29] <- "Procedencia Emergencia"
colnames(B_LabM)[38] <- "Total Facturado"
B_LabM <- B_LabM %>%select(Fecha,"Centros de Costos", 'Clasificación', Codigo,Glosa, 
                           Total, `Total Facturado`, `Procedencia Atención Abierta`, 
                           `Procedencia Atención Cerrada`, `Procedencia Emergencia`)

rm(B_Lab01,B_Lab02,B_Lab03,B_Lab04,B_Lab05,B_Lab06,B_Lab07,B_Lab08,B_Lab09,B_Lab10,B_Lab11)
B_LabM$"AT Cerrada" <- ifelse(B_LabM$`Total Facturado`==0,0,
                             B_LabM$`Total Facturado`*(B_LabM$`Procedencia Atención Cerrada`/B_LabM$Total))

# Anatomia Patologica -----------------------------------------------------

B_APM <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A1194:AL1205")
B_APM$'Clasificación' <- "Anatomia Patologica"

B_APM <- mutate_all(B_APM, ~replace(., is.na(.), 0))
B_APM$Fecha <- fecha_mes
B_APM$'Centros de Costos' <-'544-ANATOMÍA PATOLÓGICA'
colnames(B_APM)[1] <- "Codigo"
colnames(B_APM)[2] <- "Glosa"
colnames(B_APM)[3] <- "Total"
colnames(B_APM)[27] <- "Procedencia Atención Cerrada"
colnames(B_APM)[28] <- "Procedencia Atención Abierta"
colnames(B_APM)[29] <- "Procedencia Emergencia"
colnames(B_APM)[38] <- "Total Facturado"
B_APM <- B_APM %>%select(Fecha,"Centros de Costos", 'Clasificación', Codigo,Glosa, Total, `Total Facturado`, `Procedencia Atención Abierta`, `Procedencia Atención Cerrada`, `Procedencia Emergencia`)

B_APM$"AT Cerrada" <- ifelse(B_APM$`Total Facturado`==0,0,
                              B_APM$`Total Facturado`*(B_APM$`Procedencia Atención Cerrada`/B_APM$Total))


# UMT ---------------------------------------------------------------------

B_UMTM <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                   range = "B!A1164:AL1191")
B_UMTM$'Clasificación' <- "Medicina Transfusional"

B_UMTM <- mutate_all(B_UMTM, ~replace(., is.na(.), 0))
B_UMTM$Fecha <- fecha_mes
B_UMTM$'Centros de Costos' <-'575-BANCO DE SANGRE'
colnames(B_UMTM)[1] <- "Codigo"
colnames(B_UMTM)[2] <- "Glosa"
colnames(B_UMTM)[3] <- "Total"
colnames(B_UMTM)[27] <- "Procedencia Atención Cerrada"
colnames(B_UMTM)[28] <- "Procedencia Atención Abierta"
colnames(B_UMTM)[29] <- "Procedencia Emergencia"
colnames(B_UMTM)[38] <- "Total Facturado"
B_UMTM <- B_UMTM %>%select(Fecha,"Centros de Costos", 'Clasificación', Codigo,Glosa, Total, `Total Facturado`, `Procedencia Atención Abierta`, `Procedencia Atención Cerrada`, `Procedencia Emergencia`)

B_UMTM$"AT Cerrada" <- ifelse(B_UMTM$`Total Facturado`==0,0,
                             B_UMTM$`Total Facturado`*(B_UMTM$`Procedencia Atención Cerrada`/B_UMTM$Total))


# Quirofanos --------------------------------------------------------------
B_Qf_464 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B!A2005:AL2092")
B_Qf_464$'Clasificación' <- "Cirugia Cardiovascular"
B_Qf_464$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'

B_Qf_475 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A1300:AL1379")
B_Qf_475$'Clasificación' <- "Neurocirugia"
B_Qf_475$"Centros de Costos" <-'475-QUIRÓFANOS NEUROCIRUGÍA'

B_Qf_478 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A1431:AL1513")
B_Qf_478$'Clasificación' <- "Oftalmologia"
B_Qf_478$"Centros de Costos" <-'478-QUIRÓFANOS OFTALMOLOGÍA'

B_Qf_480 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A1588:AL1695")
B_Qf_480$'Clasificación' <- "Otorrino"
B_Qf_480$"Centros de Costos" <-'480-QUIRÓFANOS OTORRINOLARINGOLOGÍA'

B_Qf_493 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A1763:AL1833")
B_Qf_493$'Clasificación' <- "Cirugia Plástica"
B_Qf_493$"Centros de Costos" <-'493-QUIRÓFANOS CIRUGÍA PLÁSTICA'

B_Qf_467 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A2265:AL2377")
B_Qf_467$'Clasificación' <- "Cirugia abdominal"
B_Qf_467$"Centros de Costos" <-'467-QUIRÓFANOS DIGESTIVA'

B_Qf_486 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A2463:AL2540")
B_Qf_486$'Clasificación' <- "Urologia"
B_Qf_486$"Centros de Costos" <-'486-QUIRÓFANOS UROLOGÍA'

B_Qf_4851 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B!A2652:AL2857")
B_Qf_4851$'Clasificación' <- "Traumatologia"
B_Qf_4851$"Centros de Costos" <-'485-QUIRÓFANOS TRAUMATOLOGÍA Y ORTOPEDIA'

B_Qf_4852 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A2867:AL2869")
B_Qf_4852$'Clasificación' <- "Retiro de elementos de Osteosintesis"
B_Qf_4852$"Centros de Costos" <-'485-QUIRÓFANOS TRAUMATOLOGÍA Y ORTOPEDIA'

B_Qf_4951 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A1701:AL1759")
B_Qf_4951$'Clasificación' <- "Cabeza y cuello "
B_Qf_4951$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'

B_Qf_4952 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A1853:AL1876")
B_Qf_4952$'Clasificación' <- "Dermatologia y tegumentos"
B_Qf_4952$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'

B_Qf_4953 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A2106:AL2186")
B_Qf_4953$'Clasificación' <- "Torax"
B_Qf_4953$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'

B_Qf_4954 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A2381:AL2418")
B_Qf_4954$'Clasificación' <- "Cirugia Proctologica"
B_Qf_4954$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'

B_Qf_4955 <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                       range = "B!A2570:AL2574")
B_Qf_4955$'Clasificación' <- "Cirugia de mamas"
B_Qf_4955$"Centros de Costos" <-'464-QUIRÓFANOS CARDIOVASCULAR'


B_QfM <- rbind(B_Qf_464,B_Qf_475,B_Qf_478, B_Qf_480,B_Qf_493,B_Qf_467,B_Qf_486,B_Qf_4851,B_Qf_4852,
               B_Qf_4951,B_Qf_4952,B_Qf_4953,B_Qf_4954, B_Qf_4955)
B_QfM <- mutate_all(B_QfM, ~replace(., is.na(.), 0))
B_QfM$Fecha <- fecha_mes
colnames(B_QfM)[1] <- "Codigo"
colnames(B_QfM)[2] <- "Glosa"
colnames(B_QfM)[3] <- "Total"
colnames(B_QfM)[27] <- "Procedencia Atención Cerrada"
colnames(B_QfM)[28] <- "Procedencia Atención Abierta"
colnames(B_QfM)[29] <- "Procedencia Emergencia"
colnames(B_QfM)[38] <- "Total Facturado"
colnames(B_QfM)[15] <- "IQ mayores no ambulatorias electivas"
colnames(B_QfM)[18] <- "IQ mayores ambulatorias electivas"
colnames(B_QfM)[21] <- "IQ mayores ambulatorias de urgencia"
colnames(B_QfM)[24] <- "IQ mayores no ambulatorias de urgencia"
B_QfM <- B_QfM %>%select(Fecha,"Centros de Costos", 'Clasificación', Codigo,Glosa, Total, `Total Facturado`,
                         `Procedencia Atención Abierta`, `Procedencia Atención Cerrada`, `Procedencia Emergencia`,
                         `IQ mayores no ambulatorias electivas`,`IQ mayores ambulatorias electivas`,
                         `IQ mayores no ambulatorias de urgencia`,`IQ mayores no ambulatorias de urgencia`)
rm(B_Qf_464,B_Qf_475,B_Qf_478, B_Qf_480,B_Qf_493,B_Qf_467,B_Qf_486,B_Qf_4851,B_Qf_4852,
   B_Qf_4951,B_Qf_4952,B_Qf_4953,B_Qf_4954, B_Qf_4955) #borro los objetos innecesarios
B_QfM$"AT Cerrada" <- ifelse(B_QfM$`Total Facturado`==0,0,
                             B_QfM$`Total Facturado`*(B_QfM$`Procedencia Atención Cerrada`/B_QfM$Total))



# Quirofanos B17 ----------------------------------------------------------


B171_Qf <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                      range = "B17!B175:K188")
B172_Qf <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B17!B192:K194")
B171_Qf <- rbind(B171_Qf, B172_Qf)
B171_Qf <- mutate_all(B171_Qf, ~replace(., is.na(.), 0))
B171_Qf$Fecha <- fecha_mes
colnames(B171_Qf)[1] <- "Tipo de intervención quirúrgica"
colnames(B171_Qf)[2] <- "Total producción"
colnames(B171_Qf)[10] <- "Cirugias menores"
B171_Qf$...3 <- NULL
B171_Qf$...4 <- NULL
B171_Qf$...5 <- NULL
B171_Qf$...6 <- NULL
B171_Qf$...7 <- NULL
B171_Qf$...8 <- NULL
B171_Qf$...9 <- NULL

B172_Qf <- read_xlsx(archivoBS, na = " ",col_names = FALSE,
                     range = "B17!A199:C202")
B172_Qf$...2 <- c(B172_Qf$...1[1.1],B172_Qf$...1[2.1],"URGENCIA, MAYOR NO AMBULATORIA", "URGENCIA, MAYOR AMBULATORIA")
B172_Qf$...1 <- c("ELECTIVAS","ELECTIVAS","URGENCIA","URGENCIA")
colnames(B172_Qf)[2] <- "Tipo"
colnames(B172_Qf)[3] <- "Total"
colnames(B172_Qf)[1] <- "Ingreso"
B172_Qf$Fecha <- fecha_mes


# Lee las BBDD ------------------------------------------------------------

B_Img <- read_excel(B_IMGBBDD)
B_Lab <- read_excel(B_LABBBDD)
B_AP <- read_excel(B_APBBDD)
B_UMT <- read_excel(B_UMTBBDD)
B_Qf <- read_excel(B_QfBBDD)
B171BD_Qf <- read_excel(B171_QfBBDD)
B172BD_Qf <- read_excel(B172_QfBBDD)

# junta la informacion con el REM mensual   -------------------------------


B_Img <- rbind(B_Img, B_ImgM)
B_Lab <- rbind(B_Lab, B_LabM)
B_AP <- rbind(B_AP, B_APM)
B_UMT <- rbind(B_UMT, B_UMTM)
B_Qf <- rbind(B_Qf, B_QfM)
B_UMT <- rbind(B_UMT, B_UMTM)
B_Qf <- rbind(B_Qf, B_QfM)
B171_Qf <- rbind(B171BD_Qf, B171_Qf)
B172_Qf <- rbind(B172BD_Qf, B172_Qf)

# da formato fecha a la variable fecha ------------------------------------

B_Img$Fecha=as.Date(B_Img$Fecha)
B_Lab$Fecha=as.Date(B_Lab$Fecha)
B_AP$Fecha=as.Date(B_AP$Fecha)
B_UMT$Fecha=as.Date(B_UMT$Fecha)
B_Qf$Fecha=as.Date(B_Qf$Fecha)
B171_Qf$Fecha=as.Date(B171_Qf$Fecha)
B172_Qf$Fecha=as.Date(B172_Qf$Fecha)

# Graba las BBDD en el archivo excel --------------------------------------

openxlsx::write.xlsx(B_Img, B_IMGBBDD, colNames = TRUE, sheetName = "B_Img", overwrite = T)
openxlsx::write.xlsx(B_Lab, B_LABBBDD, colNames = TRUE, sheetName = "B_Lab", overwrite = T)
openxlsx::write.xlsx(B_AP, B_APBBDD, colNames = TRUE, sheetName = "B_AP", overwrite = T)
openxlsx::write.xlsx(B_UMT, B_UMTBBDD, colNames = TRUE, sheetName = "B_UMT", overwrite = T)
openxlsx::write.xlsx(B_Qf, B_QfBBDD, colNames = TRUE, sheetName = "B_Qf", overwrite = T)
openxlsx::write.xlsx(B171_Qf, B171_QfBBDD, colNames = TRUE, sheetName = "B171_Qf", overwrite = T)
openxlsx::write.xlsx(B172_Qf, B172_QfBBDD, colNames = TRUE, sheetName = "B172_Qf", overwrite = T)
}
