library("readxl")
library(tidyverse)
setwd(getwd())
#
StartProjectionYear <- 2022
EndProjectionYear <- 2051
FileName <- 'Model Inputs.xlsx'
user_inputs <- read_excel(FileName, sheet = 'User Inputs')
Historical_Data <- read_excel(FileName, sheet = 'Historical Data')
ScenType <- user_inputs[which(user_inputs[,1] == 'Scenario Type'),2]
#
##################################################################################################################################################################
#
RunModel <- function(user_inputs, Historical_Data, ScenType){
  #Load Inputs
  #Economic Assumptions
  dis_r <- as.double(user_inputs[which(user_inputs[,1] == 'Discount rate'),2])
  dis_r_proj <- as.double(user_inputs[which(user_inputs[,1] == 'Discount rate - Projection'),2])
  asum_infl <- as.double(user_inputs[which(user_inputs[,1] == 'Assumed Inflation'),2])
  COLA_scen <- user_inputs[which(user_inputs[,1] == 'COLA Scenario'),2]
  COLA_assum <- as.double(user_inputs[which(user_inputs[,1] == 'COLA Assumption'),2])
  #
  #Normal Cost Tiers and EE Contribution
  NC_Tier1 <- as.double(user_inputs[which(user_inputs[,1] == 'Normal Cost - Tier 1'),2])
  NC_Tier2 <- as.double(user_inputs[which(user_inputs[,1] == 'Normal Cost - Tier 2'),2])
  NC_Tier3_4 <- as.double(user_inputs[which(user_inputs[,1] == 'Normal Cost - Tier 3 & 4'),2])
  NC_Tier_NewHires <- as.double(user_inputs[which(user_inputs[,1] == 'Normal Cost (New Hires, >= 7/1/2020)'),2])
  NCAnchor_NewHire <- as.double(user_inputs[which(user_inputs[,1] == 'Anchor Normal Cost for New Hire Benefit Payments'),2])
  NCAnchor_growth <- as.double(user_inputs[which(user_inputs[,1] == 'Annual Normal Cost growth rate'),2])
  NC_StaryYear <- as.double(user_inputs[which(user_inputs[,1] == 'Start Year'),2])
  EEContrib <- as.double(user_inputs[which(user_inputs[,1] == 'Employee Contribution Rate - salary over $24,000'),2])
  EEContrib_24K <- as.double(user_inputs[which(user_inputs[,1] == 'Employee Contribution Rate - salary $24,000 or less'),2])
  EE_Over24K <- as.double(user_inputs[which(user_inputs[,1] == 'Percentage of payroll for those making more than $24,000'),2])
  #
  #ARR and Payroll
  ARR <- as.double(user_inputs[which(user_inputs[,1] == 'Actual return on assets - Projection'),2])
  Payroll_growth <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Growth Rate'),2])
  Payroll_growthT1_1 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 1'),2])
  Payroll_growthT2_1 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 1'),3])
  Payroll_growthT3_1 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 1'),4])
  Payroll_growthT1_2 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 2'),2])
  Payroll_growthT2_2 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 2'),3])
  Payroll_growthT3_2 <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Tier Growth Rate 2'),4])
  Payroll_anchor_year <- as.double(user_inputs[which(user_inputs[,1] == 'Payroll Anchor Year'),2])
  #
  #Funding Policy
  First_COLA <- as.double(user_inputs[which(user_inputs[,1] == 'First COLA'),2])
  Retirees25Y <- as.double(user_inputs[which(user_inputs[,1] == '% Retirees with over 25Yos at Retirement and below median'),2])
  Admin_Exp_Percentage <- as.double(user_inputs[which(user_inputs[,1] == 'Administrative Expenses (as % of payroll)'),2])
  ER_Policy <- user_inputs[which(user_inputs[,1] == 'Employer Contribution Policy'),2]
  ERContrib <- as.double(user_inputs[which(user_inputs[,1] == 'Statutory Employer Contribution Rate'),2])
  ER_ARP_Percentage <- as.double(user_inputs[which(user_inputs[,1] == 'Employer contribution to ERB on ARP Payroll'),2])
  NoYearsADC <- as.double(user_inputs[which(user_inputs[,1] == 'Number of Years - ADC'),2])
  Amo_Type <- user_inputs[which(user_inputs[,1] == 'Fixed/Layered Amortization'),2]
  AmoBaseInc <- as.double(user_inputs[which(user_inputs[,1] == 'Amortization Base Increase Rate'),2])
  ResetBasesZero <- as.double(user_inputs[which(user_inputs[,1] == 'Reset Bases to Zero Funding Threshold'),2])
  AVA_lowerbound <- as.double(user_inputs[which(user_inputs[,1] == 'AVA Corridor - lower bound (% of MVA)'),2])
  AVA_upperbound <- as.double(user_inputs[which(user_inputs[,1] == 'AVA Corridor - upper bound (% of MVA)'),2])
  #
  #Liability Sensitivities
  LiabSensDR <- as.double(user_inputs[which(user_inputs[,1] == 'Liability sensitivity to discount rate'),2])
  Convexity <- as.double(user_inputs[which(user_inputs[,1] == 'Convexity'),2])
  NCSensDR <- as.double(user_inputs[which(user_inputs[,1] == 'Normal cost sensitivity to discount rate'),2])
  #
  #Scenarios and Simulations
  AnalysisType <- user_inputs[which(user_inputs[,1] == 'Analysis Type'),2]
  SimType <- user_inputs[which(user_inputs[,1] == 'Simulation Type'),2]
  ContrFreeze <- user_inputs[which(user_inputs[,1] == '5 year contribution Freeze'),2]
  FitType <- user_inputs[which(user_inputs[,1] == 'Fit Type'),2]
  #
  ##################################################################################################################################################################
  #
  #Load Historical Data
  #Year and Payroll
  FYE <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'FYE')])
  TotalPayroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll Total')])
  Tier1Payroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll Tier 1 (pre-6/10)')])
  Tier2Payroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll Tier 2 (7/10-6/13)')])
  Tier3Payroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll Tier 3 (7/13-6/19)')])
  NewHirePayroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll New Hires')])
  ARPPayroll <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Payroll ARP')])
  #
  #Discount Rate, Normal Costs and Accrued Liability
  OriginalDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Original DR')])
  NewDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'New DR')])
  AccrLiabOrigDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Accrued Liability - Original DR')])
  MOYNCExistOrigDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'MOY NC Existing EEs - Original DR')])
  MOYNCNewHireOrigDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'MOY NC New Hires - Original DR')])
  AccrLiabNewDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Accrued Liability - New DR')])
  MOYNCExistNewDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'MOY NC Existing EEs - New DR')])
  MOYNCNewHireNewDR <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'MOY NC New Hires - New DR')])
  #
  #AVA and MVA
  AVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'AVA')])
  MVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'MVA')])
  ROA_MVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'ROA MVA')])
  UAL_AVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'UAL-AVA')])
  UAL_MVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'UAL-MVA')])
  FR_AVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Funded Ratio - AVA')])
  FR_MVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Funded Ratio - MVA')])
  #
  #Funding Period and Contributions
  Impl_FundPeriod <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Implied Funding Period')])
  BenPayments <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Benefit Payments')])
  AdminExp <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Admin Exp')])
  EE_Contrib <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Employee Cont')])
  ER_NC <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Employer Normal Cost')])
  ER_Amo <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Employer Amortization')])
  ER_ARP <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Employer ARP')])
  Solv_Contrib <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Solvency Contribution')])
  Total_Contrib <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Total Cont')])
  #
  #Baseline
  Years_BP <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Years of BP')])
  Total_ER <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Total Employer')])
  Total2021_ER <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Total 2021$ Employer')])
  Total_Baseline <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Total Baseline')])
  InflAdj_Baseline <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Infl Adj Baseline')])
  AVAFR_Baseline <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'AVA FR Baseline')])
  #
  #Baseline
  NetCF <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Net CF')])
  ExpInvInc <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Exp Inv Income')])
  ExpectedMVA <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Expected MVA')])
  GainLoss <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Gain / Loss')])
  DeferedCurYear <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Deferred Current Year')])
  Year1GL <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Year-1')])
  Year2GL <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Year-2')])
  Year3GL <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Year-3')])
  TotalDefered <- as.matrix(Historical_Data[,which(colnames(Historical_Data) == 'Total Deferred')])
  #
  ##################################################################################################################################################################
  #
  #Create an empty Matrix for the Projection Years
  #Get Start Index, since historical data has 3 rows, we want to start at 4
  StartIndex <- nrow(FYE) + 1
  EmptyMatrix <- matrix(0,(EndProjectionYear - StartProjectionYear + 1), 1)
  #
  FYE <- c(FYE, c(StartProjectionYear:EndProjectionYear))
  TotalPayroll <- c(TotalPayroll, EmptyMatrix)
  Tier1Payroll <- c(Tier1Payroll, EmptyMatrix)
  Tier2Payroll <- c(Tier2Payroll, EmptyMatrix)
  Tier3Payroll <- c(Tier3Payroll, EmptyMatrix)
  NewHirePayroll <- c(NewHirePayroll, EmptyMatrix)
  ARPPayroll <- c(ARPPayroll, EmptyMatrix)
  
  
  OriginalDR <- c(OriginalDR, EmptyMatrix)
  NewDR <- c(NewDR, EmptyMatrix)
  AccrLiabOrigDR <- c(AccrLiabOrigDR, EmptyMatrix)
  MOYNCExistOrigDR <- c(MOYNCExistOrigDR, EmptyMatrix)
  MOYNCNewHireOrigDR <- c(MOYNCNewHireOrigDR, EmptyMatrix)
  AccrLiabNewDR <- c(AccrLiabNewDR, EmptyMatrix)
  MOYNCExistNewDR <- c(MOYNCExistNewDR, EmptyMatrix)
  MOYNCNewHireNewDR <- c(MOYNCNewHireNewDR, EmptyMatrix)
  
  AVA <- c(AVA, EmptyMatrix)
  MVA <- c(MVA, EmptyMatrix)
  ROA_MVA <- c(ROA_MVA, EmptyMatrix)
  UAL_AVA <- c(UAL_AVA, EmptyMatrix)
  UAL_MVA <- c(UAL_MVA, EmptyMatrix)
  FR_AVA <- c(FR_AVA, EmptyMatrix)
  FR_MVA <- c(FR_MVA, EmptyMatrix)
  
  Impl_FundPeriod <- c(Impl_FundPeriod, EmptyMatrix)
  BenPayments <- c(BenPayments, EmptyMatrix)
  AdminExp <- c(AdminExp, EmptyMatrix)
  EE_Contrib <- c(EE_Contrib, EmptyMatrix)
  ER_NC <- c(ER_NC, EmptyMatrix)
  ER_Amo <- c(ER_Amo, EmptyMatrix)
  ER_ARP <- c(ER_ARP, EmptyMatrix)
  Solv_Contrib <- c(Solv_Contrib, EmptyMatrix)
  Total_Contrib <- c(Total_Contrib, EmptyMatrix)
  
  Years_BP <- c(Years_BP, EmptyMatrix)
  Total_ER <- c(Total_ER, EmptyMatrix)
  Total2021_ER <- c(Total2021_ER, EmptyMatrix)
  Total_Baseline <- c(Total_Baseline, EmptyMatrix)
  InflAdj_Baseline <- c(InflAdj_Baseline, EmptyMatrix)
  AVAFR_Baseline <- c(AVAFR_Baseline, EmptyMatrix)
  
  NetCF <- c(NetCF, EmptyMatrix)
  ExpInvInc <- c(ExpInvInc, EmptyMatrix)
  ExpectedMVA <- c(ExpectedMVA, EmptyMatrix)
  GainLoss <- c(GainLoss, EmptyMatrix)
  DeferedCurYear <- c(DeferedCurYear, EmptyMatrix)
  Year1GL <- c(Year1GL, EmptyMatrix)
  Year2GL <- c(Year2GL, EmptyMatrix)
  Year3GL <- c(Year3GL, EmptyMatrix)
  TotalDefered <- c(TotalDefered, EmptyMatrix)
  
  #Initialize Amortization and Outstnading Base
  OutstandingBase <- matrix(0,NoYearsADC, NoYearsADC + 1)
  Amortization <- matrix(0,NoYearsADC, NoYearsADC)
  AmortizationYears <- matrix(0,NoYearsADC, 1)
  #
  ##################################################################################################################################################################
  #
  #Functions
  PresentValue = function(rate, nper, pmt) {
    PV = pmt * (1 - (1 + rate) ^ (-nper)) / rate * (1 + rate)
    return(PV)
  }
  #
  ##################################################################################################################################################################
  #
  #Get ROA-MVA Index based on Scenario, this will be referenced later
  Scenario_Data <- read_excel(FileName, sheet = 'Inv_Returns')
  ScenarioIndex <- which(colnames(Scenario_Data) == as.character(ScenType))
  #
  #Benefit Payments
  BP_Data <- read_excel(FileName, sheet = 'Benefit Payments')
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
  BP_Matrix <- BP_Matrix[4:nrow(BP_Matrix),]
  #
  ##################################################################################################################################################################
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
  #Projections
  #ROA_MVA <- rnorm(31, 0.07,0.12)
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
    AdminExp[i] <- -1*Admin_Exp_Percentage*TotalPayroll[i]
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
      EE_Contrib[i]<- EE_Contrib[i-1]
      ER_NC[i] <- ER_NC[i-1]
      ER_ARP[i] <- ER_ARP[i-1]
      ER_Amo[i] <- ER_Amo[i-1]
    } else {
      EE_Contrib[i] <- TotalPayroll[i]*(EE_Over24K*(EEContrib - EEContrib_24K) + EEContrib_24K)
      ER_NC[i] <- MOYNCExistNewDR[i-1] + MOYNCNewHireNewDR[i-1] - AdminExp[i] - EE_Contrib[i]
      ER_ARP[i] <- ARPPayroll[i]*ER_ARP_Percentage
      
      if(ER_Policy == 'Statutory Rate'){
        ER_Amo[i] <- ERContrib*TotalPayroll[i] - ER_NC[i]
      } else {
        if(count > nrow(Amortization)){
          ER_Amo[i] <- max(0 - ER_ARP[i], -(ER_NC[i] + ER_ARP[i]))
        } else {
          ER_Amo[i] <- max(sum(Amortization[count,]) - ER_ARP[i], -(ER_NC[i] + ER_ARP[i]))
        }
      }
    }
    CashFlows <- (BenPayments[i] + AdminExp[i] + EE_Contrib[i] + ER_NC[i] + ER_ARP[i] + ER_Amo[i])
    ROA_MVA[i] <- as.double(Scenario_Data[count,ScenarioIndex])
    Solv_Contrib[i] <- max(-(MVA[i-1]*(1+ROA_MVA[i]) + CashFlows*(1+ROA_MVA[i])^0.5) / (1+ROA_MVA[i])^0.5,0)
    Total_ER[i] <- ER_NC[i] + ER_Amo[i] + ER_ARP[i] + Solv_Contrib[i]
    Total2021_ER[i] <<- Total_ER[i] / (((1 + asum_infl)^(FYE[i] - NC_StaryYear))*TotalPayroll[i])
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
    UAL_AVA[i] <<- AccrLiabNewDR[i] - AVA[i]
    FR_AVA[i] <<- AVA[i]/AccrLiabNewDR[i]
    UAL_MVA[i] <- AccrLiabNewDR[i] - MVA[i]
    FR_MVA[i] <- MVA[i]/AccrLiabNewDR[i]
    #
    #Amo Policy, Outstanding Balance
    #Amo Years
    if(Amo_Type == 'Layered'){
      AmortizationYears[count+1] <- NoYearsADC
    } else {
      AmortizationYears[count+1] <- max(AmortizationYears[count+1] - 1,0)
    }
    
    if(count < nrow(Amortization)){
      #Oustanding Balance
      for(j in 2:(count + 1)){
        OutstandingBase[count+1,j] <- OutstandingBase[count,j-1]*(1 + NewDR[i-1]) - (Amortization[count,j-1]*(1 + NewDR[i-1])^0.5)
      }
      OutstandingBase[count+1,1] <- AccrLiabNewDR[i] - AVA[i] - sum(OutstandingBase[count+1,2:ncol(OutstandingBase)])
      
      #Amo Layers
      for(j in 1:(count + 1)){
        rate <- ((1 + NewDR[i]) / (1 + AmoBaseInc)) - 1
        period <- max(AmortizationYears[count+1,1] - j + 1, 1)
        pmt <- 1
        Amortization[count+1,j] <- OutstandingBase[count+1,j] / (PresentValue(rate,period,pmt) / ((1+NewDR[i-1])^0.5))
      }
    }
  }
  Output <- cbind(FYE,TotalPayroll,Tier1Payroll,Tier2Payroll,Tier3Payroll,NewHirePayroll,ARPPayroll,OriginalDR,NewDR,AccrLiabOrigDR,MOYNCExistOrigDR,MOYNCNewHireNewDR)
  Output <- cbind(Output,AccrLiabNewDR,MOYNCExistNewDR,MOYNCNewHireNewDR,BenPayments,AdminExp,EE_Contrib,ER_NC,ER_ARP,ER_Amo,Solv_Contrib,Total_ER,Total2021_ER)
  Output <- cbind(Output,NetCF,ExpInvInc,ExpectedMVA,GainLoss,DeferedCurYear,Year1GL,Year2GL,Year3GL,TotalDefered,MVA,AVA,UAL_AVA,UAL_MVA,FR_AVA,FR_MVA)
}

#Scenarios <- c('Assumption','Model','Recession','Recurring Recession')
#Initialize Matrix for Scenarios
#Scenario_UAL <- as.data.frame(FYE)
#Scenario_FR <- as.data.frame(FYE)
#Scenario_ER <- as.data.frame(FYE)

#for (i in 1:length(Scenarios)){
#   NewScenario <- Scenarios[i]
#   RunModel(user_inputs, Historical_Data, NewScenario)
#   
#   Scenario_UAL <- cbind(Scenario_UAL,UAL_AVA)
#   Scenario_FR <- cbind(Scenario_FR,FR_AVA)
#   Scenario_ER <- cbind(Scenario_ER,Total2021_ER)
#}

#for (i in 1:10000){
#  FR_AVA <- RunModel(user_inputs, Historical_Data, 'Assumption')
#  Scenario_FR <- cbind(Scenario_FR,FR_AVA)
#}
#colnames(Scenario_UAL) <- c('FYE',Scenarios)
#colnames(Scenario_FR) <- c('FYE',Scenarios)
#write_excel_csv(as.data.frame(Output),'Model Outputs.csv')