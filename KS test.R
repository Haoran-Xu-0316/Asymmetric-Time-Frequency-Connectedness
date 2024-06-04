library(readxl)
library(openxlsx)


ntotal=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_npdc.xlsx','Total')
nt13=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_npdc.xlsx','1-3')
nt36=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_npdc.xlsx','3-6')
nt6=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_npdc.xlsx','6')

ptotal=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_npdc.xlsx','Total')
pt13=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_npdc.xlsx','1-3')
pt36=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_npdc.xlsx','3-6')
pt6=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_npdc.xlsx','6')



# for (i in 2:16) {print(ks.test(as.numeric(ntotal[[i]]), as.numeric(ptotal[[i]])))}



t_ks_results <- data.frame(statistic = numeric(), p_value = numeric())
t13_ks_results <- data.frame(statistic = numeric(), p_value = numeric())
t36_ks_results <- data.frame(statistic = numeric(), p_value = numeric())
t6_ks_results <- data.frame(statistic = numeric(), p_value = numeric())


for (i in 2:16) {
  
  result <- ks.test(as.numeric(ntotal[[i]]), as.numeric(ptotal[[i]]))
  t_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}

for (i in 2:16) {
  
  result <- ks.test(as.numeric(nt13[[i]]), as.numeric(pt13[[i]]))
  t13_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}

for (i in 2:16) {
  
  result <- ks.test(as.numeric(nt36[[i]]), as.numeric(pt36[[i]]))
  t36_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}

for (i in 2:16) {

  result <- ks.test(as.numeric(nt6[[i]]), as.numeric(pt6[[i]]))
  t6_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}




wb <- createWorkbook()

addWorksheet(wb, "t")
addWorksheet(wb, "13")
addWorksheet(wb, "36")
addWorksheet(wb, "6")



writeData(wb, "t", t_ks_results)
writeData(wb, "13", t13_ks_results)
writeData(wb, "36", t36_ks_results)
writeData(wb, "6", t6_ks_results)


saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\NPDC_Test_Results.xlsx", overwrite = TRUE)





########################################################################################
#######################################################################################
######################################################################################



NTo=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_to.xlsx','to')
PTo=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_to.xlsx','to')


NFrom=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_from.xlsx','from')
PFrom=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_from.xlsx','from')



NNet=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_net.xlsx','net')
PNet=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_net.xlsx','net')





TO_ks_results <- data.frame(statistic = numeric(), p_value = numeric())
FROM_ks_results <- data.frame(statistic = numeric(), p_value = numeric())
NET_ks_results <- data.frame(statistic = numeric(), p_value = numeric())




for (i in 2:25) {
  
  result <- ks.test(as.numeric(NTo[[i]]), as.numeric(PTo[[i]]))
  TO_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}

for (i in 2:25) {
  
  result <- ks.test(as.numeric(NFrom[[i]]), as.numeric(PFrom[[i]]))
  FROM_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}

for (i in 2:25) {
  
  result <- ks.test(as.numeric(NNet[[i]]), as.numeric(PNet[[i]]))
  NET_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}





wb <- createWorkbook()

addWorksheet(wb, "TO")
addWorksheet(wb, "FROM")
addWorksheet(wb, "NET")



writeData(wb, "TO", TO_ks_results)
writeData(wb, "FROM", FROM_ks_results)
writeData(wb, "NET", NET_ks_results)


saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\TFN_Test_Results.xlsx", overwrite = TRUE)



########################################################################################
#######################################################################################
######################################################################################



NTCI=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Negative_total.xlsx','total')
PTCI=read_excel('C:\\Users\\徐浩然\\Desktop\\P2 Results\\Positive_total.xlsx','total')




TCI_ks_results <- data.frame(statistic = numeric(), p_value = numeric())





for (i in 2:5) {
  
  result <- ks.test(as.numeric(NTCI[[i]]), as.numeric(PTCI[[i]]))
  TCI_ks_results[i-1, ] <- c(result$statistic, result$p.value)
}





wb <- createWorkbook()

addWorksheet(wb, "TCI")



writeData(wb, "TCI", TCI_ks_results)



saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\TCI_Test_Results.xlsx", overwrite = TRUE)



