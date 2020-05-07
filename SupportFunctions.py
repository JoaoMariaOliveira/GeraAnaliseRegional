
import numpy as np
import pandas as pd
import string
import xlrd
import numpy.core.multiarray
from pandas import ExcelWriter
from pandas import ExcelFile
# ============================================================================================
def Calc_MultiplierI(vVar, nStatesSectors, vTotalProduction, mLeontief):
    vVar=vVar.reshape(1, nStatesSectors)
    vCoef = np.zeros([1, nStatesSectors], dtype=float)
    for j in range(nStatesSectors):
        vCoef[0, j] = vVar[0, j] / vTotalProduction[0, j]

    vGerador = np.zeros([1,nStatesSectors], dtype=float)
    for j in range(nStatesSectors):
        vGerador[0,j] = np.sum(mLeontief[:,j] * vCoef[0,:])

    vMultiplier = np.zeros([1, nStatesSectors], dtype=float)
    vMultiplier[0, :] = vGerador[0, :] / vCoef[0, :]
    vMultiplier = np.nan_to_num(vMultiplier, nan=0, posinf=0, neginf=0)
    return vMultiplier.reshape(nStatesSectors,1), vGerador.reshape(nStatesSectors,1), vCoef.reshape(nStatesSectors,1)

# ============================================================================================
def Calc_MultiplierII(vCoef, nStatesSectors, mLeontief):
    vCoef=vCoef.reshape(1, nStatesSectors)
    vGerador = np.zeros([1,nStatesSectors], dtype=float)
    for j in range(nStatesSectors):
        vGerador[0,j] = np.sum(mLeontief[0:nStatesSectors,j] * vCoef[0,:])

    vMultiplier = np.zeros([1, nStatesSectors], dtype=float)
    vMultiplier[0,:] = vGerador[0, :] / vCoef[0, :]
    vMultiplier = np.nan_to_num(vMultiplier, nan=0, posinf=0, neginf=0)
    return vMultiplier.reshape(nStatesSectors,1), vGerador.reshape(nStatesSectors,1)

# ============================================================================================
def read_file_excel(sDirectory, sFileName, sSheetName):
    sFileName = sDirectory + sFileName
    mSheet = pd.read_excel(sFileName, sheet_name=sSheetName, header=None, index_col=None)
    return mSheet

# =============================================================
# Função que lê dados de um arquivo txt e retorna uma matriz
# lê de uma pasta input
# ============================================================================================

def read_file_txt(sFileName, sDirectory):
    sFileName = sDirectory + sFileName
    temp = np.loadtxt(sFileName, dtype=np.float64, skiprows= 1)
    return temp

# ============================================================================================
def load_data_comextstat(sDirectoryInput, sFile, nStates, nProducts):
    mSheet = read_file_txt(sFile, sDirectoryInput)
#    mSheet = read_file_excel(sDirectoryInput, sFile, sSheet)
    mValue = np.zeros([nStates, nProducts], dtype=float)
    for s in range(nStates):
        for p in range(nProducts):
            mValue[s, p] =mSheet[nProducts*s+p, 5]
    return mValue

# ============================================================================================
def load_table_states(sDirectoryInput, sFileStates, nStates):
    mSheet = read_file_excel(sDirectoryInput, sFileStates, 'TabStates')
    vCodStates = []
    vNameRegions = []
    vNameStates = []
    vShortNameStates = []
    vRegionState = []
    for s in range(nStates):
        vCodStates.append(mSheet.values[s+1, 0])
#        vNameRegions.append(mSheet.values[s+1, 1])
        vRegionState.append(mSheet.values[s+1, 2])
        vNameStates.append(mSheet.values[s+1, 3])
        vShortNameStates.append(mSheet.values[s+1, 4])

    return vCodStates, vRegionState, vNameStates, vShortNameStates
# ============================================================================================
# ============================================================================================
def load_convert_states(sDirectoryInput, sFileStates, nStates):
    mSheet = read_file_excel(sDirectoryInput, sFileStates, 'Convert')
    nLinIni=29
    nColIni=1

    vConvertStates = np.zeros([nStates], dtype=int)
    for l in range(nStates):
        sAux=mSheet.values[nLinIni+l, nColIni+l]
        vConvertStates[l]=mSheet.values[nLinIni-1, nColIni+l]

    return vConvertStates
# ============================================================================================
def load_DemandShock(sDirectoryInput, sFileDemandShock, sSheetName, nSectors, nAdjustUni):
    mSheet = pd.read_excel(sDirectoryInput + sFileDemandShock, sheet_name=sSheetName, header=None)
    nColIni = 2
    nLinIni = 1
    mDemandShock=np.zeros([nSectors,6],dtype=float)
    for i in range(nSectors):
        mDemandShock[i,:] = nAdjustUni + mSheet.values[nLinIni + i, nColIni:nColIni+6]

    return mDemandShock
# ============================================================================================
def load_OfferShock(sDirectoryInput, sFileDemandShock, sSheetName, nSectors, nAdjustUni):
    mSheet = pd.read_excel(sDirectoryInput + sFileDemandShock, sheet_name=sSheetName, header=None)
    nColIni = 2
    nLinIni = 1
    vLaborRestriction=np.zeros([nSectors,1],dtype=float)
    for i in range(nSectors):
        vLaborRestriction[i,0]= nAdjustUni + mSheet.values[nLinIni + i, nColIni]

    return vLaborRestriction
# ============================================================================================

def load_GrupSector(sDirectoryInput, sFileInput, nGrupSectors, nSectors):
    mSheet1=read_file_excel(sDirectoryInput, sFileInput, "SetorGrupo")
    mSheet2=read_file_excel(sDirectoryInput, sFileInput, "Grupo18")
    vNameGrupSector=[]
    vCodGrupSector=np.zeros([nSectors], dtype=int)
    for i in range(nSectors):
        vCodGrupSector[i]=mSheet1.values[1+i, 3]

    for i in range(nGrupSectors):
        vNameGrupSector.append(mSheet2.values[1 + i, 1])

    return vCodGrupSector, vNameGrupSector
# ============================================================================================
def load_CabecalhoMIP(sDirectoryInput, sFileInput, sSheetNational, nSector ):
    mSheet=pd.read_excel(sDirectoryInput+sFileInput, header =None)
    nColIni = 3
    nLinIni = 4
    nColsDemandNat=7
    vCodSector = []
    vNameSector= []
    vNameTaxes = []
    vNameVA    = []
    vNameDemand = []

    for l in range(nSector):
        vCodSector.append(mSheet.values[nLinIni + l, 0])
        vNameSector.append(mSheet.values[nLinIni + l,1])

    for c in range(nColsDemandNat):
        nCAdd = nColIni + nSector + 1 + c
        vNameDemand.append(mSheet.values[2,nCAdd])

    for l in range(7):
        nLAdd=nLinIni + nSector + 2 + l
        vNameTaxes.append(mSheet.values[nLAdd, 1])
    for l in range(15):
        nLAdd = nLinIni + nSector + 10 + l
        vNameVA.append(mSheet.values[nLAdd, 1])

    return vCodSector, vNameSector, vNameTaxes, vNameVA, vNameDemand
# ============================================================================================
def load_RegionalMIP(sDirectoryInput, sFileInput, sSheet, nStates, nSector):
    mSheet=pd.read_excel(sDirectoryInput+sFileInput, sheet_name=sSheet, header=None)
    mMatrix = mSheet.values
    nColsDemandReg=6
    nColIni = 2
    nLinIni = 3
    nSectorStates= nStates * nSector
    nDemandStates =nStates * nColsDemandReg
    mMipAux = mMatrix[nLinIni:nSectorStates+25+nLinIni, nColIni:nSectorStates+1+nDemandStates+nColIni+2]

    # Regional - Intermediate consum  + National Production +Imports + taxes + Intermediate Consum Total Line + Added Value
    mIntermConsumReg=np.copy(mMipAux[0:nSectorStates,0:nSectorStates])
    vNatProductReg=np.copy(mMipAux[nSectorStates,0:nSectorStates]).reshape(1,nSectorStates)
    vImportsReg =np.copy(mMipAux[nSectorStates+1,0:nSectorStates]).reshape(1,nSectorStates)
    mTaxesReg=np.copy(mMipAux[nSectorStates+2:nSectorStates+9,0:nSectorStates])
    vIntermConsumTotProdReg=np.copy(mMipAux[nSectorStates+9,0:nSectorStates]).reshape(1,nSectorStates)
    mVAReg =np.copy(mMipAux[nSectorStates+10:nSectorStates+25,0:nSectorStates])
    # Intermediate Consum Total Column + Intermediate Consum Total Taxes  Column+ Intermediate Consum Total Tazes Column
    vIntermConsumTotConsReg=np.copy(mMipAux[0:nSectorStates,nSectorStates]).reshape(nSectorStates,1)
    vProdIntermConsumNatTotConsumReg = np.copy(mMipAux[nSectorStates, nSectorStates]).reshape(1, 1)
    vImportsConsumNatTotReg = np.copy(mMipAux[nSectorStates+1, nSectorStates]).reshape(1, 1)
    vTotTaxesIntermConsumReg=np.copy(mMipAux[nSectorStates+2:nSectorStates+9,nSectorStates]).reshape(7,1)
    vIntermConsumTotConsumTotReg = np.copy(mMipAux[nSectorStates+9, nSectorStates]).reshape(1, 1)
    vTotVAIntermConsumReg=np.copy(mMipAux[nSectorStates+10:nSectorStates+25,nSectorStates]).reshape(15,1)
    # Regional - Demad + total national Demand Line + Imports Demand + Taxes Demand + Total taxes Demand line + Added Value Demand
    mDemandReg=np.copy(mMipAux[0:nSectorStates, nSectorStates+1:nSectorStates+1+nDemandStates])
    vNatDemandReg=np.copy(mMipAux[nSectorStates, nSectorStates+1:nSectorStates+1+nDemandStates]).reshape(1,nDemandStates)
    vImportsDemandReg=np.copy(mMipAux[nSectorStates+1, nSectorStates+1:nSectorStates+1+nDemandStates]).reshape(1,nDemandStates)
    mTaxesDemandReg=np.copy(mMipAux[nSectorStates+2:nSectorStates+9, nSectorStates+1:nSectorStates+1+nDemandStates])
    vTotDemandReg=np.copy(mMipAux[nSectorStates+9, nSectorStates+1:nSectorStates+1+nDemandStates]).reshape(1,nDemandStates)
    mVADemandReg =np.copy(mMipAux[nSectorStates+10:nSectorStates+25, nSectorStates+1:nSectorStates+1+nDemandStates])
    # Regional - Final Demand + total Regional Final Demand Point + Taxes Regional Final Demand + Total Final Demand point +   Added Value Final Demand
    vTotFinalDemandReg=np.copy(mMipAux[0:nSectorStates, nSectorStates+1+nDemandStates]).reshape(nSectorStates,1)
    vTotProdNacFinalDemandReg=np.copy(mMipAux[nSectorStates, nSectorStates+1+nDemandStates]).reshape(1,1)
    vTotImportsFinalDemandReg=np.copy(mMipAux[nSectorStates+1, nSectorStates+1+nDemandStates]).reshape(1,1)
    vTotTaxesFinalDemandReg=np.copy(mMipAux[nSectorStates+2:nSectorStates+9, nSectorStates+1+nDemandStates]).reshape(7,1)
    vTotIntermConsumFinalDemandReg = np.copy(mMipAux[nSectorStates+9, nSectorStates+1+nDemandStates]).reshape(1, 1)
    vTotVAFinalDemandReg=np.copy(mMipAux[nSectorStates+10:nSectorStates+25, nSectorStates+1+nDemandStates]).reshape(15,1)
    # regional - Total Demand + total national Demand Point + Taxes total Demand + Total Demand point +   Added Value Total Demand
    vTotalDemandReg=np.copy(mMipAux[0:nSectorStates, nSectorStates+1+nDemandStates+1]).reshape(nSectorStates,1)
    vTotProdNacDemandReg = np.copy(mMipAux[nSectorStates, nSectorStates+1+nDemandStates+1]).reshape(1, 1)
    vTotImportsDemandReg=np.copy(mMipAux[nSectorStates+1, nSectorStates+1+nDemandStates+1]).reshape(1,1)
    vTotalTaxesReg=np.copy(mMipAux[nSectorStates+2:nSectorStates+9, nSectorStates+1+nDemandStates+1]).reshape(7,1)
    vTotIntermConsumDemandReg = np.copy(mMipAux[nSectorStates+9, nSectorStates+1+nDemandStates+1]).reshape(1, 1)
    vTotalVAReg=np.copy(mMipAux[nSectorStates+10:nSectorStates+25, nSectorStates+1+nDemandStates+1]).reshape(15,1)

    return mIntermConsumReg, vNatProductReg, vImportsReg, mTaxesReg, vIntermConsumTotProdReg, mVAReg, \
           vIntermConsumTotConsReg, vProdIntermConsumNatTotConsumReg, vImportsConsumNatTotReg, vTotTaxesIntermConsumReg, vIntermConsumTotConsumTotReg, vTotVAIntermConsumReg,\
           mDemandReg, vNatDemandReg, vImportsDemandReg, mTaxesDemandReg, vTotDemandReg, mVADemandReg,\
           vTotFinalDemandReg, vTotProdNacFinalDemandReg, vTotImportsFinalDemandReg, vTotTaxesFinalDemandReg, vTotVAFinalDemandReg, vTotIntermConsumFinalDemandReg, \
           vTotalDemandReg, vTotProdNacDemandReg, vTotImportsDemandReg, vTotalTaxesReg, vTotIntermConsumDemandReg, vTotalVAReg
# ============================================================================================
# Função que grava dados em um arquivo excel
# Grava em uma pasta output e pode gravar várias planilhas em um mesmo arquivo
# ============================================================================================

def write_data_excel(sDirectoryOutput, sFileName, lSheetName, lDataSheet, lRowsLabel, lColsLabel, vUseHeader):
    Writer = pd.ExcelWriter(sDirectoryOutput+sFileName, engine='xlsxwriter')
    df=[]
    for each in range(len(lSheetName)):
        if vUseHeader[each] == True:
            df.append(pd.DataFrame(lDataSheet[each],  index=lRowsLabel[each], columns=lColsLabel[each], dtype=float))
            df[each].to_excel(Writer, lSheetName[each], header=True, index=True)
        else:
            df.append(pd.DataFrame(lDataSheet[each],  dtype=float))
            df[each].to_excel(Writer, lSheetName[each], header=False, index=False)

    Writer.save()

# ============================================================================================
# Função que grava dados em um arquivo excel
# Grava em uma pasta output e pode gravar uma só planilha em um mesmo arquivo
# ============================================================================================

def write_file_excel(sDirectoryOutput, sFileName, sSheet, mData, vRows, vCols, lUseHeader):
    Writer = pd.ExcelWriter(sDirectoryOutput+sFileName, engine='xlsxwriter')
    if lUseHeader == True:
        df=pd.DataFrame(mData, index=vRows, columns=vCols, dtype=float)
        df.to_excel(Writer, sheet_name=sSheet, header=True, index=True)
    else:
        df=pd.DataFrame(mData, dtype=float)
        df.to_excel(Writer, sheet_name=sSheet, header=False, index=False)
    Writer.save()

# ============================================================================================