library(foreign)
df=read.spss("C:/Users/User/Downloads/Data.sav", to.data.frame = TRUE)
View(df)
attach(df)

library(Publish)
sex=sutable(Sex~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
type=sutable(TypeoftransplantAuto1Allo2~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
Weight=sutable(Weightgrouplt30kg1gt30kg2~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
TBV=sutable(TBVGROUPlt20001gt20002~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)

Blood=sutable(BloodpriminggroupYes1No2~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)

TBV_processed=sutable(TBVGROUPlt20001gt20002~Q(ACDused)+Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)


## Exporting multiple excel sheet in one excel file
library(writexl) 
sheets <- list("sheet1Name" = sex, "sheet2Name" =type,"sheet3Name"=Weight,"sheet4Name"=TBV,"sheet5Name"=Blood,"sheet6Name"=TBV_processed) 
#assume sheet1 and sheet2 are data frames
write_xlsx(sheets,"C:/Users/User/Downloads/pradip.xlsx")

##########################################################


adverse=sutable(OverallAdverseeventsUNEVENTFUL1REACTION2~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
citrate=sutable(CitratereactionsAE1~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
AE2=sutable(AccessrelatedAE2~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
AE3=sutable( GIAE3~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
AE4=sutable(TechnicalAE4~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
AE5=sutable(OthersAE5~Q(productvolume)+Q(ProductMNC)+Q(ProductTLC)+Q(productCD34),data=df)
  

## Exporting multiple excel sheet in one excel file
library(writexl) 
sheets <- list("sheet1Name" = adverse, "sheet2Name" =citrate,"sheet3Name"=AE2,"sheet4Name"=AE3,"sheet5Name"=AE4,"sheet6Name"=AE5) 
#assume sheet1 and sheet2 are data frames
write_xlsx(sheets,"C:/Users/User/Downloads/pradip1.xlsx")
