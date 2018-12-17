library('qcc')
library('readxl')
library('ggplot2')
library('moments')
library('openxlsx')
library('EnvStats')
library('ppcc')
library('sn')
library('MASS')
library('fitdistrplus')

#stops packages from causing the "Hit <return> to see next plot:
devAskNewPage(ask = FALSE)


total = 100



Blisk_Compiled_Data <- read_excel("Blisk Compiled Data Modifieds.xlsx", 
                                  col_types = c("numeric", "text", "text", 
                                                "date", "numeric", "numeric", "numeric", 
                                                "numeric", "numeric"))

data = as.data.frame(Blisk_Compiled_Data)



rm(Blisk_Compiled_Data)
#Specs for the different blisk parameters
blisks = unique(data$Part)
specs = c(5,7,4,56,12)

cp1 = list()
cpk1 = list()
skew1 = list()
kurt1 = list()
parts1 = list()
parameter1 = list()
originalppcc1 = list()
modifiedppcc1 = list()
specdens1 = list()
cnames1 = list()

specprob1 = list()


counter = 0
datapoints = 6000

probablilitydata = matrix(nrow = datapoints,ncol = length(blisks)* length(specs)*3)


for (val in blisks){
  
  #select only some blisks to analyze

  
  
  
  datatoprocess <- data[ which(data$Part==val),]

  #determines which blisks to analyze
  last = length(datatoprocess$Part)
  first = last - total
  #datatoprocess = datatoprocess[first:last,]
  
  
  for(i in 1:length(specs)){
    #stops packages from causing the "Hit <return> to see next plot:
    devAskNewPage(ask = FALSE)
    
    counter = counter + 1
    columnname = colnames( datatoprocess)[4 + i] 
    
    
    #xmat is just a matrix of 1's
    #data is fitted and then a density line is plotted
    xmat = matrix(1, nrow = dim(datatoprocess)[1], ncol = 1)
    
    distfit = selm.fit(x = xmat, y = datatoprocess[,4 + i])
    
    f1  = makeSECdistr(dp= distfit$param$dp, family="SN", name="First-SN")
    

    scalarforspec =0.5
    
    test2 <- plot(f1, npt = datapoints, range= c(0,specs[i] + specs[i]*scalarforspec), probs = c(0.95, 0.99865, 0.999968, 0.9999997), xlab = (expression(paste(mu,'"'))),main = paste(columnname," Surface roughness for Part ",val))
    #adds vertical line for spec
    abline(v = specs[i],col = "red")

    #Saves plots as PDF's
    #dev.copy(pdf,paste("Process Capability Density Plot",columnname,val,".pdf"))
    #dev.off()
    
    
    probloc1 = test2[["x"]]
    probdens1 = test2[["density"]]
    specloc = probloc1[1/(1 + scalarforspec)*6000]
    
    
    specdens1[counter] = probdens1[specloc]
    
    probablilitydata[,3*counter - 2] = probloc1
    probablilitydata[,3*counter - 1] = probdens1
    probablilitydata[,3*counter] = rep.int(0,length(probdens1))
    
    
    cnames1[3*counter - 2] = paste(columnname,val)
    cnames1[3*counter - 1] = paste(columnname,val)
    cnames1[3*counter] = "Total Probability"
    
    
    
    

    #Typical process capability charting with xbar, R chart and process capability
    
    
    xbarplot <- qcc(datatoprocess[,4 + i], type="xbar.one",data.name = paste(columnname,val), labels = datatoprocess[,4])
    xbarplot <- qcc(modified, type="xbar.one",data.name = paste(columnname,val), labels = datatoprocess[,4])
    axes.las = 3
    #Saves plots as PDF's
    dev.copy(pdf,paste("X Bar Plot",columnname,val,".pdf"))
    dev.off()
    
    processcapability <- process.capability(xbarplot,spec.limits=c(0,specs[i]))
    processcapability <- process.capability(xbarplot,spec.limits=c(0,specs[i]))
    dev.copy(pdf,paste("Process Capability",columnname,val,".pdf"))
    dev.off()
    
    creates range data
    rData <- matrix(cbind(datatoprocess[1:dim(datatoprocess)[1]-1,4 + i], datatoprocess[2:dim(datatoprocess)[1],4 + i]), ncol=2)
    rData <- matrix(cbind(modified[1:length(modified)-1], modified[2:length(modified)]), ncol=2)
    
    rbarplot <- qcc(rData,type = "R",data.name = paste(columnname,val))
    dev.copy(pdf,paste("Moving Range",columnname,val,".pdf"))
    dev.off()
    
    

    
    cp1[counter] = processcapability[["indices"]][1,1]
    cpk1[counter] = processcapability[["indices"]][3,1]
    parts1[counter] = val
    parameter1[counter] = columnname
    
    #kurtosis and skewness of data
    skew1[counter] = skewness(datatoprocess[,4 + i])
    kurt1[counter] = kurtosis(datatoprocess[,4 + i])
  }
}
parts = unlist(parts1)
parameter = unlist(parameter1)


specprob = unlist(specprob1)

specdata = data.frame(parts,parameter,specprob)

write.xlsx(specdata, 'Non-conformity of parts at Spec.xlsx')


modifiedppcc = unlist(modifiedppcc1)
originalppcc = unlist(originalppcc1)



cp = unlist(cp1)
cpk = unlist(cpk1)
skew = unlist(skew1)
kurt = unlist(kurt1)
specdens = unlist(specdens1)
cnames = unlist(cnames1)
processcapabilties = data.frame(parts,parameter, cp, cpk,skew,kurt,specdens)
processspecs = data.frame(parts,parameter,specdens)
colnames(probablilitydata) = cnames
Parts =as.character(parts)


write.xlsx(processspecs, 'Process Spec Capabilities.xlsx')
write.xlsx(probablilitydata, 'Probability Data.xlsx')



# cp for processes
ggplot(processcapabilties, aes(fill = Parts, y=cp, x=parameter)) + 
  geom_bar(position="dodge", stat="identity") +
  ggtitle("cp for Processes") +
  theme(plot.title = element_text(hjust = 0.5)) +
  geom_hline(yintercept= 2, linetype="dashed", color = "red")

# cpk for processes
ggplot(processcapabilties, aes(fill=Parts, y=cpk, x=parameter)) + 
  geom_bar(position="dodge", stat="identity") +
  ggtitle("cpk for Processes") +
  theme(plot.title = element_text(hjust = 0.5)) +
  geom_hline(yintercept= 2, linetype="dashed", color = "red")

# Skewness for processes
ggplot(processcapabilties, aes(fill=Parts, y=skew, x=parameter)) + 
  geom_bar(position="dodge", stat="identity") +
  ggtitle("Skewness for Processes") +
  theme(plot.title = element_text(hjust = 0.5)) +
  geom_hline(yintercept= 2, linetype="dashed", color = "red")

# Kurtosis for processes
ggplot(processcapabilties, aes(fill=Parts, y=kurt, x=parameter)) + 
  geom_bar(position="dodge", stat="identity") +
  ggtitle("Kurtosis for Processes") +
  theme(plot.title = element_text(hjust = 0.5)) +
  geom_hline(yintercept= 3, linetype="dashed", color = "red")

