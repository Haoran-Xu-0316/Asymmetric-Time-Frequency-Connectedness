require(xts)
library(urca)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)
library(zoo)
library(knitr)
library(FinTS)
library(vars)
library(readxl)
library(tseries)
library(parallel)
library(DistributionUtils)
library(ConnectednessApproach)

label=c('EUA',	'SZA',	'HBA', 'PCM', 'CEI', 'GBI')


t=as.list(label)
tt=list()
for (i in 1:length(label)) {
  if (i+1 <= length(label)) {
    for (j in (i+1):length(label)){
      # print(i)
      # print(j)
      tt=append(tt, paste(t[i],"-",t[j]))
      
    }}}
tt <- c("Date", tt)



read_data = function(path, sheet) {
  
  data = read_excel(path, sheet = sheet)
  dates <- as.Date(data$Date, format = "%Y%m%d")
  data <- xts(data[, label], order.by = dates)
  data <- na.approx(data)
  rt=data
  return(rt)
  
}




DATA=read_data('C:\\Users\\徐浩然\\Desktop\\P2Data_1M.xlsx','C2_p')
DATA=as.data.frame(DATA)
dt=as.zoo(DATA, order.by = as.Date(rownames(DATA)))
VARselect(DATA,lag.max=20,type="const") 



cor(DATA,  method =  "kendall")


Y = Yp = Yn = dt[-1,]
k = ncol(Y)
for (i in 1:k) {
  x = embed(as.numeric(dt[,i]),2)
  Y[,i] = Yp[,i] = Yn[,i] = 100*log(x[,1]/x[,2])
  Yp[which(Y[,i]<0),i] = 0
  Yn[which(Y[,i]>0),i] = 0
}
Y_list = list(Y, Yp, Yn)


partition = c(pi+0.00001, pi/5, pi/22, 0)
DCA = list()
spec = c("all", "positive", "negative")
for (i in 1:length(Y_list)) {
  DCA[[i]] = suppressMessages(ConnectednessApproach(Y_list[[i]], 
                                                    model="TVP-VAR",
                                                    connectedness="Frequency",
                                                    nlag=1,
                                                    nfore=20,
                                                    # window.size=200,
                                                    VAR_config=list(TVPVAR=list(kappa1=0.99, kappa2=0.99, prior="BayesPrior")),
                                                    Connectedness_config = list(FrequencyConnectedness=list(partition=partition, generalized=TRUE, scenario="ABS"))))
  kable(DCA[[i]]$TABLE)
}





PlotTCI(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotFROM(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotTO(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotNET(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotNPDC(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))

PlotPCI(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotINF(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotNetwork(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))
PlotNPT(DCA[[1]], ca=list(DCA[[2]], DCA[[3]]))




DATA=as.data.frame(DATA)
df= row.names(DATA)
# 


####################################################################################################
## 保存结果
####################################################################################################


Svexl= function(dca,txt){
  
  
  Retr <- function(df, n, matrix_type) {
    for (i in 1:length(label)) {
      if (i + 1 <= length(label)) {
        for (j in (i + 1):length(label)) {
          if (matrix_type == "inf") {
            r <- as.numeric(inf[i, j, , n])
          } else if (matrix_type == "pci") {
            r <- as.numeric(pci[i, j, , n])
          } else if (matrix_type == "npdc") {
            r <- as.numeric(npdc[i, j, , n])
          }
          df <- cbind(df, r)
        }
      }
    }
    names(df) <- tt
    return(df)
  }
  
table=dca$TABLE
total=dca$TCI
from=dca$FROM
to=dca$TO
net=dca$NET
npdc=dca$NPDC
pci=dca$PCI
inf=dca$INFLUENCE






##导出数据
npdct=Retr(df,1,'npdc')
npdc13=Retr(df,2,'npdc')
npdc36=Retr(df,3,'npdc')
npdc6=Retr(df,4,'npdc')

pcit=Retr(df,1,'pci')
pci13=Retr(df,2,'pci')
pci36=Retr(df,3,'pci')
pci6=Retr(df,4,'pci')



inft=Retr(df,1,'inf')
inf13=Retr(df,2,'inf')
inf36=Retr(df,3,'inf')
inf6=Retr(df,4,'inf')




total=as.data.frame(total)
to=as.data.frame(to)
from=as.data.frame(from)
net=as.data.frame(net)
npdc=as.data.frame(npdc)




wb1 <- createWorkbook()

addWorksheet(wb1, "Total")
writeData(wb1, "Total", npdct[-1,])
addWorksheet(wb1, "1-3")
writeData(wb1, "1-3", npdc13[-1,])
addWorksheet(wb1, "3-6")
writeData(wb1, "3-6", npdc36[-1,])
addWorksheet(wb1, "6")
writeData(wb1, "6", npdc6[-1,])

saveWorkbook(wb1, paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_npdc.xlsx"), overwrite = TRUE)



wb2 <- createWorkbook()

addWorksheet(wb2, "Total")
writeData(wb2, "Total", pcit[-1,])
addWorksheet(wb2, "1-3")
writeData(wb2, "1-3", pci13[-1,])
addWorksheet(wb2, "3-6")
writeData(wb2, "3-6", pci36[-1,])
addWorksheet(wb2, "6")
writeData(wb2, "6", pci6[-1,])

saveWorkbook(wb2, paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_pci.xlsx"), overwrite = TRUE)


wb3 <- createWorkbook()

addWorksheet(wb3, "Total")
writeData(wb3, "Total", inft[-1,])
addWorksheet(wb3, "1-3")
writeData(wb3, "1-3", inf13[-1,])
addWorksheet(wb3, "3-6")
writeData(wb3, "3-6", inf36[-1,])
addWorksheet(wb3, "6")
writeData(wb3, "6", inf6[-1,])

saveWorkbook(wb3, paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_inf.xlsx"), overwrite = TRUE)
# 
write.csv(table,file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_table.csv")) ##关联矩阵
write.xlsx(total, file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_total.xlsx"), sheetName = 'total', rowNames = TRUE)  ## 总溢出指数
write.xlsx(to, file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_to.xlsx"), sheetName = 'to', rowNames = TRUE)  ## 溢出指数
write.xlsx(from, file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_from.xlsx"), sheetName = 'from', rowNames = TRUE)  ## 溢入指数
write.xlsx(net, file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", txt, "_net.xlsx"), sheetName = 'net', rowNames = TRUE)  ## 净溢出

}

Svexl(DCA[[1]],'All')

Svexl(DCA[[2]],'Positive')

Svexl(DCA[[3]],'Negative')


####################################################################################################
## 保存结果
####################################################################################################



####################################################################################################
## 对正负收益率都进行descriptive statistics
####################################################################################################

stat=function(data,id){
  
  Series     = colnames(data)
  N          = length(Series)
  stats      = matrix(0,14,N,dimnames=list(c('Mean','Max.','Min.','Std. Dev.','Skewness','Kurtosis','J-B', 'p-value' ,'ERS','p-value','L-B','p-value','L-B2','p-value'),Series))
  
  for (i in 1:N){
    
    stats[1,i]  = mean(data[,Series[i]])
    stats[2,i]  = max(data[,Series[i]])
    stats[3,i]  = min(data[,Series[i]])
    stats[4,i]  = sd(data[,Series[i]])
    stats[5,i]  = skewness(data[,Series[i]])
    stats[6,i]  = kurtosis(data[,Series[i]])
    
    stats[7,i]  = jarque.bera.test(data[,Series[i]])$statistic
    stats[8,i]  = jarque.bera.test(data[,Series[i]])$p.value
    stats[9,i] = ur.ers(data[,Series[i]], type="DF-GLS", model="const", lag.max=0)@testreg[["coefficients"]][1,3]
    stats[10,i] = ur.ers(data[,Series[i]], type="DF-GLS", model="const", lag.max=0)@testreg[["coefficients"]][1,4]
    stats[11,i] = Box.test(data[,Series[i]],10,"Ljung-Box")$statistic
    stats[12,i] = Box.test(data[,Series[i]],10, "Ljung-Box")$p.value
    stats[13,i] = Box.test(data[,Series[i]]^2,10,"Ljung-Box")$statistic
    stats[14,i] = Box.test(data[,Series[i]]^2,10, "Ljung-Box")$p.value
  }
  print(t(stats), digits=6)
  write.csv(stats, file = paste0("C:\\Users\\徐浩然\\Desktop\\P2 Results\\", id, "_Descriptive.csv"))
}

stat(Y,"Y")
stat(Yp,"Yp")
stat(Yn,"Yn")


####################################################################################################
## 对正负收益率都进行descriptive statistics
####################################################################################################



