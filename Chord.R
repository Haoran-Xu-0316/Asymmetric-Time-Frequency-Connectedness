# Libraries
library(tidyverse)
library(dplyr)
library(readxl)
library(viridis)
library(patchwork)
library(hrbrthemes)
library(circlize)
library(chorddiag)  #devtools::install_github("mattflor/chorddiag")
windowsFonts(myFont = windowsFont("Palatino Linotype"))


l=as.list(c(
"(a) Normal. Total",
"(b) Normal. 1-5",
"(c) Normal. 5-22",
"(d) Normal. 22-inf",
"(e) Positive. Total",
"(f) Positive. 1-5",
"(g) Positive. 5-22",
"(h) Positive. 22-inf",
"(i) Negative. Total",
"(j) Negative. 1-5",
"(k) Negative. 5-22",
"(l) Negative. 22-inf"
))


data1 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "all_t")
data1=data1[,-1]
rownames(data1) <- colnames(data1)
data_long1 <- data1 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data2 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "all_15")
data2=data2[,-1]
rownames(data2) <- colnames(data2)
data_long2 <- data2 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data3 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "all_522")
data3=data3[,-1]
rownames(data3) <- colnames(data3)
data_long3 <- data3 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data4 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "all_22")
data4=data4[,-1]
rownames(data4) <- colnames(data4)
data_long4 <- data4 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data5 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "p_t")
data5=data5[,-1]
rownames(data5) <- colnames(data5)
data_long5 <- data5 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data6 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "p_15")
data6=data6[,-1]
rownames(data6) <- colnames(data6)
data_long6 <- data6 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data7 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "p_522")
data7=data7[,-1]
rownames(data7) <- colnames(data7)
data_long7 <- data7 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data8 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "p_22")
data8=data8[,-1]
rownames(data8) <- colnames(data8)
data_long8 <- data8 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)

data9 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "n_t")
data9=data9[,-1]
rownames(data9) <- colnames(data9)
data_long9 <- data9 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data10 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "n_15")
data10=data10[,-1]
rownames(data10) <- colnames(data10)
data_long10 <- data10 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data11 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "n_522")
data11=data11[,-1]
rownames(data11) <- colnames(data11)
data_long11 <- data11 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)


data12 <- read_excel("C:\\Users\\徐浩然\\Desktop\\NPDCchord.xlsx",sheet = "n_22")
data12=data12[,-1]
rownames(data12) <- colnames(data12)
data_long12 <- data12 %>%
  rownames_to_column %>%
  gather(key = 'key', value = 'value', -rowname)



# parameters
# 
circos.clear()


mycolorppp=c(CSI ="#ADB5C2", GPR= "#6E6F4D", VIX = "#98A070",EPU="#C1BB9F",'S&P' ="#E4B576", DJSI = "#F6DFC0",  BEA = "#DCA8A4",EUA= "#C06B5F")

# color palette

mycolor=c(EUA ="#E4B576", SZA=  "#DCA8A4", HBA = "#98A070",PMI="#ADB5C2",CEI ="#F6DFC0", GBI ="#6E6F4D")
names(mycolor)=colnames(data)
# EUA
# SZA
# HBA
# PMI
# CEI
# GBI

# png("C:\\Users\\徐浩然\\Desktop\\Chord.png",res = 80*8, height = 900*8,width = 600*8)
svg("C:\\Users\\徐浩然\\Desktop\\Paper 2 Figure\\Chord.svg", height = 8.3, width = 10)
layout(matrix(1:12, 3, 4,byrow = TRUE))

for (i in 1:12){
  if (i ==1){data_long=data_long1}
  if (i ==2){data_long=data_long2}
  if (i ==3){data_long=data_long3}
  if (i ==4){data_long=data_long4}
  if (i ==5){data_long=data_long5}
  if (i ==6){data_long=data_long6}
  if (i ==7){data_long=data_long7}
  if (i ==8){data_long=data_long8}
  if (i ==9){data_long=data_long9}
  if (i ==10){data_long=data_long10}
  if (i ==11){data_long=data_long11}
  if (i ==12){data_long=data_long12}

  
par(mar = rep(0, 4) ,family="Palatino Linotype")
circos.par(start.degree = 90, gap.degree = 4, 
           track.margin = c(-0.08, 0.1),
           points.overflow.warning = FALSE)
# Base plot
chordDiagram(
  x = data_long, 
  grid.col = mycolor,
  # row.col = mycolor,
  transparency = 0.25,
  directional = 1,
  direction.type = c("diffHeight","arrows"),
  diffHeight  = 0.04,#可正负

  annotationTrack = "grid",
  annotationTrackHeight = c(0.05, 0.1),
  link.arr.type = "big.arrow",
  link.sort = TRUE,
  # link.target.prop=TRUE,
  target.prop.height=0.02,
  # reduce=-1,
  link.arr.length = 0.04,
  link.largest.ontop = FALSE,
  
  
  
  
  # link.lwd = 1,    # Line width
  # link.lty = 1,    # Line type
  # link.border = 1 # Border color
  )

# Add text and axis
circos.trackPlotRegion(
  track.index = 1,
  bg.border = NA,
  panel.fun = function(x, y) {

    xlim = get.cell.meta.data("xlim")
    sector.index = get.cell.meta.data("sector.index")

    # Add names to the sector.
    circos.text(
      x = mean(xlim),
      y = 4,
      labels = sector.index,
      facing = "bending",
      cex = 0.9,font=2
    )

    # Add graduation on axis
    circos.axis(
      h = "top",
      labels.cex = 0.6,
      # major.at = seq(from = 0, to = xlim[2], by = ifelse(test = xlim[2]>10, yes = 4, no = 2)),
      minor.ticks = 1,
      major.tick.length = 0.3,
      labels.niceFacing = FALSE)
    # cex.axis = 1
  }
  )
text(0,-1.12, l[i],cex=1.2,font=2)
circos.clear()
}
dev.off()
