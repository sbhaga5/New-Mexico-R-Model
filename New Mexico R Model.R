library("readxl")
library(ggplot2)
library(tidyverse)
setwd(getwd())
rm(list = ls())
#
#User can change these values
StartYear <- 2018
StartProjectionYear <- 2021
EndProjectionYear <- 2051
FileName <- 'Model Inputs.xlsx'
#
#Reading Input File
user_inputs_numeric <- read_excel(FileName, sheet = 'Numeric Inputs')
user_inputs_character <- read_excel(FileName, sheet = 'Character Inputs')
Historical_Data <- read_excel(FileName, sheet = 'Historical Data')
Scenario_Data <- read_excel(FileName, sheet = 'Inv_Returns')
#
##################################################################################################################################################################
#
#Function for Present Value for Amortization
PresentValue = function(rate, nper, pmt) {
  PV = pmt * (1 - (1 + rate) ^ (-nper)) / rate * (1 + rate)
  return(PV)
}

#Assigning numeric inputs
for(i in 1:nrow(user_inputs_numeric)){
  if(!is.na(user_inputs_numeric[i,2])){
    assign(as.character(user_inputs_numeric[i,2]),as.double(user_inputs_numeric[i,3]), envir = .GlobalEnv)
  }
}

#Assigning character inputs
for(i in 1:nrow(user_inputs_character)){
  if(!is.na(user_inputs_character[i,2])){
    assign(as.character(user_inputs_character[i,2]),as.character(user_inputs_character[i,3]), envir = .GlobalEnv)
  }
}

#Create an empty Matrix for the Projection Years
EmptyMatrix <- matrix(0,(EndProjectionYear - StartProjectionYear + 1), 1)
for(j in 1:length(colnames(Historical_Data))){
  TempMatrix <- rbind(as.matrix(Historical_Data[,j]), EmptyMatrix)
  assign(as.character(colnames(Historical_Data)[j]), TempMatrix, envir = .GlobalEnv)
}
#Assign values for Projection Years
FYE <- StartYear:EndProjectionYear
#Get Start Index, since historical data has 3 rows, we want to start at 4
StartIndex <- StartProjectionYear - StartYear + 1

if(SimType == 'Assumed'){
  SimReturn <- SimReturnAssumed
} else if(AnalysisType == 'Conservative'){
  SimReturn <- SimReturnConservative
}

#Initialize Amortization and Outstnading Base
OutstandingBase <- matrix(0,NoYearsADC, NoYearsADC + 1)
Amortization <- matrix(0,NoYearsADC, NoYearsADC)
AmortizationYears <- matrix(0,NoYearsADC, 1)

BP_Data <- read_excel(FileName, sheet = 'Benefit Payments')
AmoYearsInput <- read_excel(FileName, sheet = 'Amortization')
#
##################################################################################################################################################################
#
#Benefit Payments
BP_Matrix <- matrix(0, nrow = (EndProjectionYear-StartProjectionYear+4), ncol = (120-60+1))
for(i in 1:nrow(BP_Matrix)){
  for(j in 1:ncol(BP_Matrix)){
    AgeFromBP <- as.double(BP_Data[1,j+1])
    SurvivalProb <- as.double(BP_Data[2,j+1])
    if((i == 1) || (j == 1)) {
      BP_Matrix[i,j] <- as.double(BP_Data[i+3,j+1])
    } 
    
    if ((j > 1) && (i > 1)){
      if(AgeFromBP >= First_COLA){
        BP_COLA <- COLA_assum
      } else {
        BP_COLA <- 0
      }
      BP_Matrix[i,j] <- BP_Matrix[i-1,j-1]*SurvivalProb*(1+BP_COLA)
    }
  }
}
#Remove First 2, keep only projection years
BP_Matrix <- BP_Matrix[3:nrow(BP_Matrix),]
#
#Offset Matrix
OffsetYears <- matrix(0,NoYearsADC, NoYearsADC)
for(i in 1:nrow(OffsetYears)){
  RowCount <- i
  ColCount <- 1
  #This is to create the "zig-zag" pattern of amo years. you also want it to become 0 if the amo years is greater than the payments
  #Meaning if its 10 years, then after 10 years, the amo payment is 0
  while(RowCount <= nrow(OffsetYears) && (ColCount <= as.double(AmoYearsInput[i,2]))){
    OffsetYears[RowCount,ColCount] <- as.double(AmoYearsInput[i,2])
    RowCount <- RowCount + 1
    ColCount <- ColCount + 1
  }
}
#
#Initialize first row of Amortization and Outstanding Base
AmortizationYears[1] <- NoYearsADC
OutstandingBase[1,1] <- AccrLiabNewDR[3] - AVA[3] 
rate <- ((1 + NewDR[3]) / (1 + AmoBaseInc)) - 1
period <- AmortizationYears[1,1]
pmt <- 1
Amortization[1,1] <- OutstandingBase[1,1] / (PresentValue(rate,period,pmt) / ((1+NewDR[3])^0.5))
#
##################################################################################################################################################################
#
RunModel <- function(user_inputs, Historical_Data, Scenario_Data, AnalysisType, SimType, SimReturn, SimVolatility, ScenType, BP_Matrix, AmoYearsInput, Amortization, OutstandingBase, OffsetYears){
  #Projections
  ScenarioIndex <- which(colnames(Scenario_Data) == as.character(ScenType))
  for(i in StartIndex:length(FYE)){
    #Payroll
    TotalPayroll[i] <- TotalPayroll[i-1]*(1 + Payroll_growth)
    PayrollTier1Growth <- (1 - Payroll_growthT1_1) - (Payroll_growthT1_1*Payroll_growthT1_2)*(FYE[i] - Payroll_anchor_year)
    PayrollTier2Growth <- (1 - Payroll_growthT2_1) - (Payroll_growthT2_1*Payroll_growthT2_2)*(FYE[i]- Payroll_anchor_year)
    PayrollTier3Growth <- (1 - Payroll_growthT3_1) - (Payroll_growthT3_1*Payroll_growthT3_2)*(FYE[i] - Payroll_anchor_year)
    Tier1Payroll[i] <- Tier1Payroll[i-1]*PayrollTier1Growth
    Tier2Payroll[i] <- Tier2Payroll[i-1]*PayrollTier2Growth
    Tier3Payroll[i] <- Tier3Payroll[i-1]*PayrollTier3Growth
    NewHirePayroll[i] <- TotalPayroll[i] - (Tier1Payroll[i] + Tier2Payroll[i] + Tier3Payroll[i])
    ARPPayroll[i] <- ARPPayroll[i-1]*(TotalPayroll[i] / TotalPayroll[i-1])
    #
    #Discount Rate
    OriginalDR[i] <- dis_r
    NewDR[i] <- dis_r_proj
    #
    #Benefit Payments, Admin Expenses
    count <- i - StartIndex + 1
    BenPayments[i] <- -1*sum(BP_Matrix[count,])
    AdminExp[i] <- -1*Admin_Exp_Percentage*(TotalPayroll[i])
    #
    #Accrued Liability, MOY NC - Original DR
    if(i > (StartIndex)){
      NCGrowth <- (1 + NCAnchor_growth)^(FYE[i]  - NC_StaryYear + 1)
      MOYNCExistOrigDR[i-1] <- (Tier1Payroll[i]*(NC_Tier1) + Tier2Payroll[i]*(NC_Tier2) + Tier3Payroll[i]*(NC_Tier3_4))*NCGrowth
      MOYNCNewHireOrigDR[i-1] <- NewHirePayroll[i]*(NCAnchor_NewHire)*NCGrowth
    }
    AccrLiabOrigDR[i] <- AccrLiabOrigDR[i-1]*(1+OriginalDR[i-1]) + (MOYNCExistOrigDR[i-1] + MOYNCNewHireOrigDR[i-1] + BenPayments[i])*(1+OriginalDR[i-1])^0.5
    #
    #Accrued Liability, MOY NC - New DR
    DRDifference <- 100*(OriginalDR[i]-NewDR[i])
    if(i > (StartIndex)){
      MOYNCExistNewDR[i-1] <- MOYNCExistOrigDR[i-1]*(1+(NCSensDR/100))^(DRDifference)
      MOYNCNewHireNewDR[i-1] <- MOYNCNewHireOrigDR[i-1]*(1+(NCSensDR/100))^(DRDifference) 
    }
    AccrLiabNewDR[i] <- AccrLiabOrigDR[i]*((1+(LiabSensDR/100))^(DRDifference))*((1+(Convexity/100))^((DRDifference)^2/2))
    #
    #Contribution Policy, Amortization Policy
    if((FYE[i] < 2026) && (ContrFreeze == 'FREEZE')){
      EE_Contrib[i] <- EE_Contrib[i-1]
      ER_NC[i] <- ER_NC[i-1]
      ER_ARP[i] <- ER_ARP[i-1]
      ER_Amo[i] <- ER_Amo[i-1]
    } else {
      EE_Contrib[i] <- (TotalPayroll[i])*(EE_Over24K*(EEContrib - EEContrib_24K) + EEContrib_24K)
      ER_NC[i] <- MOYNCExistNewDR[i-1] + MOYNCNewHireNewDR[i-1] - AdminExp[i] - EE_Contrib[i]
      ER_ARP[i] <- ARPPayroll[i]*ER_ARP_Percentage
      
      if(ER_Policy == 'Statutory Rate'){
        ER_Amo[i] <- ERContrib*(TotalPayroll[i]) - ER_NC[i]
      } else {
        if(count > nrow(Amortization)){
          ER_Amo[i] <- max(0 - ER_ARP[i], -(ER_NC[i] + ER_ARP[i]))
        } else {
          ER_Amo[i] <- max(sum(Amortization[count,]) - ER_ARP[i], -(ER_NC[i] + ER_ARP[i]))
        }
      }
    }
    CashFlows <- (BenPayments[i] + AdminExp[i] + EE_Contrib[i] + ER_NC[i] + ER_ARP[i] + ER_Amo[i])
    #
    #Return data based on deterministic or stochastic
    if(AnalysisType == 'Stochastic'){
      ROA_MVA[i] <- rnorm(1,SimReturn,SimVolatility)
    } else if(AnalysisType == 'Deterministic'){
      ROA_MVA[i] <- as.double(Scenario_Data[count,ScenarioIndex]) 
    }
    #
    #Solvency Contribution and Employer Contribution
    Solv_Contrib[i] <- max(-(MVA[i-1]*(1+ROA_MVA[i]) + CashFlows*(1+ROA_MVA[i])^0.5) / (1+ROA_MVA[i])^0.5,0)
    Total_ER[i] <- ER_NC[i] + ER_Amo[i] + ER_ARP[i] + Solv_Contrib[i]
    Total2021_ER[i] <- Total_ER[i] / ((1 + asum_infl)^(FYE[i] - NC_StaryYear))
    Total2021_ER_Percentage[i] <- Total_ER[i] / TotalPayroll[i]
    #
    #Net CF, Expected MVA
    NetCF[i] <- CashFlows + Solv_Contrib[i]
    ExpInvInc[i] <- (MVA[i-1]*NewDR[i-1]) + (NetCF[i]*NewDR[i-1]*0.5)
    ExpectedMVA[i] <- MVA[i-1] + NetCF[i] + ExpInvInc[i]
    MVA[i] <- MVA[i-1]*(1+ROA_MVA[i]) + NetCF[i]*(1+ROA_MVA[i])^0.5
    #
    #Gain Loss, Defered Losses
    GainLoss[i] <- MVA[i] - ExpectedMVA[i] 
    DeferedCurYear[i] <- GainLoss[i]*(0.8/1)
    Year1GL[i] <- DeferedCurYear[i-1]*(0.6/0.8)
    Year2GL[i] <- Year1GL[i-1]*(0.4/0.6)
    Year3GL[i] <- Year2GL[i-1]*(0.2/0.4)
    TotalDefered[i] <- Year1GL[i] + Year2GL[i] + Year3GL[i]
    #
    AVA_Temp <- as.double(MVA[i] - TotalDefered[i])
    AVA_Temp_1 <- as.double(MVA[i])
    AVA[i] <- min(max(AVA_Temp, AVA_Temp_1*AVA_lowerbound), AVA_Temp_1*AVA_upperbound)
    UAL_AVA[i] <- AccrLiabNewDR[i] - AVA[i]
    FR_AVA[i] <- AVA[i]/AccrLiabNewDR[i]
    UAL_MVA[i] <- AccrLiabNewDR[i] - MVA[i]
    FR_MVA[i] <- MVA[i]/AccrLiabNewDR[i]
    
    if(count < nrow(Amortization)){
      #Oustanding Balance
      for(j in 2:(count + 1)){
        OutstandingBase[count+1,j] <- OutstandingBase[count,j-1]*(1 + NewDR[i-1]) - (Amortization[count,j-1]*(1 + NewDR[i-1])^0.5)
      }
      OutstandingBase[count+1,1] <- AccrLiabNewDR[i] - AVA[i] - sum(OutstandingBase[count+1,2:ncol(OutstandingBase)])
      
      #Amo Layers
      for(j in 1:(count + 1)){
        rate <- ((1 + NewDR[i]) / (1 + AmoBaseInc)) - 1
        period <- max(OffsetYears[count+1,j] - j + 1, 1)
        pmt <- 1
        if(period > 0){
          Amortization[count+1,j] <- OutstandingBase[count+1,j] / (PresentValue(rate,period,pmt) / ((1+NewDR[i-1])^0.5))
        }
      }
    }
  }
  
  Output <- cbind(FYE,TotalPayroll,Tier1Payroll,Tier2Payroll,Tier3Payroll,NewHirePayroll,ARPPayroll,OriginalDR,NewDR,AccrLiabOrigDR,MOYNCExistOrigDR,MOYNCNewHireNewDR)
  Output <- cbind(Output,AccrLiabNewDR,MOYNCExistNewDR,MOYNCNewHireNewDR,BenPayments,AdminExp,EE_Contrib,ER_NC,ER_ARP,ER_Amo,Solv_Contrib,Total_ER,Total2021_ER)
  Output <- cbind(Output,NetCF,ExpInvInc,ExpectedMVA,GainLoss,DeferedCurYear,Year1GL,Year2GL,Year3GL,TotalDefered,ROA_MVA, MVA,AVA,UAL_AVA,UAL_MVA,FR_AVA,FR_MVA)
  
  TempData <- cbind(ROA_MVA,UAL_AVA, FR_AVA, Total2021_ER_Percentage)
  return(TempData)
}
#
##################################################################################################################################################################
#
RunModel(user_inputs, Historical_Data, Scenario_Data, 'Deterministic','', SimReturn, SimVolatility, 'Assumption', BP_Matrix, AmoYearsInput, Amortization, OutstandingBase, OffsetYears)

#Scenarios
Scenarios <- c('Assumption','Model','Recession','Recurring Recession')
#Initialize Matrix for Scenarios
Scenario_Returns <- as.data.frame(FYE)
Scenario_UAL <- as.data.frame(FYE)
Scenario_FR <- as.data.frame(FYE)
Scenario_ER <- as.data.frame(FYE)

for (i in 1:length(Scenarios)){
  NewScenario <- Scenarios[i]
  NewData <- RunModel(user_inputs, Historical_Data, Scenario_Data, 'Deterministic','', SimReturn, SimVolatility, NewScenario, BP_Matrix, AmoYearsInput, Amortization, OutstandingBase, OffsetYears)
  
  Scenario_Returns <- cbind(Scenario_Returns,NewData[,1])
  Scenario_UAL <- cbind(Scenario_UAL,NewData[,2])
  Scenario_FR <- cbind(Scenario_FR,NewData[,3])
  Scenario_ER <- cbind(Scenario_ER,NewData[,4])
}

colnames(Scenario_Returns) <- c('FYE',Scenarios)
colnames(Scenario_UAL) <- c('FYE',Scenarios)
colnames(Scenario_FR) <- c('FYE',Scenarios)
colnames(Scenario_ER) <- c('FYE',Scenarios)

ScenarioPlot <- function(Data){
  ggplot(Data, aes(x = Data$FYE)) +
    geom_line(aes(y = Data$Assumption), color = "#FF6633", size = 2) +
    geom_line(aes(y = Data$Model), color = "#FFCC33", size = 2) +
    geom_line(aes(y = Data$Recession), color = "#0066CC", size = 2) +
    geom_line(aes(y = Data$`Recurring Recession`), color = "#CC0000", size = 2)
}
ScenarioPlot(Scenario_ER)
#
##################################################################################################################################################################

# #Simulations
# start_time <- Sys.time()
# 
# NumberofSimulations <- 10000
# #RandomReturns <- replicate(NumberofSimulations, rnorm(33, SimReturn,SimVolatility))
# Returns_Sims <- matrix(1:length(FYE),nrow = length(FYE), ncol = NumberofSimulations + 1)
# UAL_Sims <- matrix(1:length(FYE),nrow = length(FYE), ncol = NumberofSimulations + 1)
# FR_Sims <- matrix(1:length(FYE),nrow = length(FYE), ncol = NumberofSimulations + 1)
# ER_Sims <- matrix(1:length(FYE),nrow = length(FYE), ncol = NumberofSimulations + 1)
# for (i in 1:NumberofSimulations){
#   NewData <- RunModel(user_inputs, Historical_Data, Scenario_Data, 'Stochastic','Assumed', SimReturn, SimVolatility, '', BP_Matrix, AmoYearsInput, Amortization, OutstandingBase, OffsetYears)
#   Returns_Sims[,i+1] <- NewData[,1]
#   UAL_Sims[,i+1] <- NewData[,2]
#   FR_Sims[,i+1] <- NewData[,3]
#   ER_Sims[,i+1] <- NewData[,4]
# }
# 
# Simulations_Returns <- cbind(FYE,FYE,FYE)
# Simulations_UAL <- cbind(FYE,FYE,FYE)
# Simulations_FR <- cbind(FYE,FYE,FYE)
# Simulations_ER <- cbind(FYE,FYE,FYE)
# 
# for(i in 1:length(FYE)){
#   Simulations_Returns[i,] <- t(quantile(Returns_Sims[i,2:ncol(Returns_Sims)],c(0.25,0.5,0.75)))
#   Simulations_UAL[i,] <- t(quantile(UAL_Sims[i,2:ncol(UAL_Sims)],c(0.25,0.5,0.75)))
#   Simulations_FR[i,] <- t(as.data.frame(quantile(FR_Sims[i,2:ncol(FR_Sims)],c(0.25,0.5,0.75))))
#   Simulations_ER[i,] <- t(quantile(ER_Sims[i,2:ncol(ER_Sims)],c(0.25,0.5,0.75)))
# }
# 
# SimulationPlot <- function(Data, FYE){
#   Data <- (as.data.frame(Data))
#   Data <- cbind(FYE, Data)
#   colnames(Data) <- c('FYE','25th Percentile', '50th Percentile', '75th Percentile')
#   ggplot(Data, aes(x = Data[,1])) +
#     geom_line(aes(y = Data[,2]), color = "#FF6633", size = 2) +
#     geom_line(aes(y = Data[,3]), color = "#FFCC33", size = 2) +
#     geom_line(aes(y = Data[,4]), color = "#0066CC", size = 2)
# }
# SimulationPlot(Simulations_FR, FYE)
# 
# end_time <- Sys.time()
# print(end_time - start_time)
