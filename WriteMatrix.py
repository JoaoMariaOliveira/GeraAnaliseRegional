# ============================================================================================
# Main Module of Building Input output Regional
# Coded to Python by João Maria de Oliveira -
# ============================================================================================
import SupportFunctions as Support
import numpy as np

def WriteCoefsMatrix(mA, mABarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy()
    vNameRowsReg[1] = vNameSectorReg[1].copy()
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy()
    vNameColsReg[1] = vNameSectorReg[1].copy()
    vDataSheet.append(mA)
    vSheetName.append('CoefsTec')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy() + ["Remunerações"]
    vNameRowsReg[1] = vNameSectorReg[1].copy() + [" "]
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy() + ['Consumo']
    vNameColsReg[1] = vNameSectorReg[1].copy() + [" "]

    vDataSheet.append(mABarr)
    vSheetName.append('CoefsTecBarr')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    sFileSheet = 'CoefsTec_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteLeontief(mLeontief, mLeontiefBarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy()
    vNameRowsReg[1] = vNameSectorReg[1].copy()
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy()
    vNameColsReg[1] = vNameSectorReg[1].copy()
    vDataSheet.append(mLeontief)
    vSheetName.append('Leontief')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy() + ["Remunerações"]
    vNameRowsReg[1] = vNameSectorReg[1].copy() + [" "]
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy() + ['Consumo']
    vNameColsReg[1] = vNameSectorReg[1].copy() + [" "]

    vDataSheet.append(mLeontiefBarr)
    vSheetName.append('LeontiefBarr')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    sFileSheet = 'Leontief_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteGhosh(mGhosh, mGhoshBarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy()
    vNameRowsReg[1] = vNameSectorReg[1].copy()
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy()
    vNameColsReg[1] = vNameSectorReg[1].copy()
    vDataSheet.append(mGhosh)
    vSheetName.append('Ghosh')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy() + ["Remunerações"]
    vNameRowsReg[1] = vNameSectorReg[1].copy() + [" "]
    vNameColsReg = [[], []]
    vNameColsReg[0] = vNameSectorReg[0].copy() + ['Consumo']
    vNameColsReg[1] = vNameSectorReg[1].copy() + [" "]

    vDataSheet.append(mGhoshBarr)
    vSheetName.append('GhoshfBarr')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    sFileSheet = 'Ghosh_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)

    return

def WriteMultipliers(tTuplesIa, tTuplesIb, tTuplesII, vNameSectorReg, sDirectoryOutput, nYear, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = [[], []]
    vNameRowsReg[0] = vNameSectorReg[0].copy()
    vNameRowsReg[1] = vNameSectorReg[1].copy()

    vNameColsReg = ['VA_I', 'VA_II', 'Occup_I', 'Occup_II', 'Remun_I', 'Remun_II', 'EOB_I', 'EOB_II', 'Salary_I',
                    'Salary_II']
    vNameMult = ['Multiplier', 'Gerador', 'Coef']

    xI = np.concatenate(tTuplesIa, axis=1)
    vDataSheet.append(xI)
    vSheetName.append(vNameMult[0])
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    xI = np.concatenate(tTuplesIb, axis=1)
    vDataSheet.append(xI)
    vSheetName.append(vNameMult[1])
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vNameColsReg = ['VA', 'Occup', 'Remun', 'EOB', 'Salary']
    xI = np.concatenate(tTuplesII, axis=1)
    vDataSheet.append(xI)
    vSheetName.append(vNameMult[2])
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    sFileSheet = 'Multip_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return
# ============================================================================================
# writing  Intensitys Matrix
# ============================================================================================
def WriteIntensityMatrix(mIntensity, vShortNameStates, vNameSector, sDirectoryOutput, nYear, nStates, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = vNameSector[0:nSectors]
    vNameColsReg = vNameSector[0:nSectors]
    for s in range(nStates):
        vDataSheet.append(mIntensity[s,:,:])
        vSheetName.append('MatInt_'+vShortNameStates[s])
        vRowsLabel.append(vNameRowsReg)
        vColsLabel.append(vNameColsReg)
        vUseHeader.append(True)

    sFileSheet = 'MatIntens_' + str(nYear) + '_' + str(nSectors)+'.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return
# ============================================================================================
# writing  Intensitys Matrix
# ============================================================================================
def WriteIntensityMatrixGrouped(mIntensity,  vShortNameStates, vNameGrupSector, sDirectoryOutput, nYear, nStates, nGrupSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = vNameGrupSector[0:nGrupSectors]
    vNameColsReg = vNameGrupSector[0:nGrupSectors]
    for s in range(nStates):
        vDataSheet.append(mIntensity[s,:,:])
        vSheetName.append('MatInt_'+vShortNameStates[s])
        vRowsLabel.append(vNameRowsReg)
        vColsLabel.append(vNameColsReg)
        vUseHeader.append(True)

    sFileSheet = 'MatIntens_' + str(nYear) + '_' + str(nGrupSectors)+'.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

# ============================================================================================
# writing  multiplier of production
# ============================================================================================
def WriteMultiplierProduction(vProdMultLeontiefI, vProdMultLeontiefII, vShortNameStates, vNameSector, sDirectoryOutput,
                              nYear, nStates, nSectors, nStatesSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = vNameSector[0:nSectors]
    vDataSheet.append(vProdMultLeontiefI.reshape(nStates, nSectors))
    vSheetName.append('MultProdLeontief_I')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = vNameSector[0:nSectors]
    vDataSheet.append((vProdMultLeontiefII[0, 0:nStatesSectors]).reshape(nStates, nSectors))
    vSheetName.append('MultProdLeontief_II')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    sFileSheet = 'MultipProd_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return
# ============================================================================================
# writing  multiplier of production
# ============================================================================================
def WriteLinkages(vBackLinkL, vForwLinkG, vBackLinkG, vForwLinkL, vShortNameStates, vNameSector, sDirectoryOutput, nYear, nStates, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = vNameSector[0:nSectors]
    vDataSheet.append(vBackLinkL.reshape(nStates,nSectors))
    vSheetName.append('BackLinkL')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vDataSheet.append(vForwLinkG.reshape(nStates,nSectors))
    vSheetName.append('ForwLinkG')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vDataSheet.append(vBackLinkG.reshape(nStates, nSectors))
    vSheetName.append('BackLinkG')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vDataSheet.append(vForwLinkL.reshape(nStates,nSectors))
    vSheetName.append('ForwLinkL')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    sFileSheet = 'Linkage_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteIntraInter(mProdMultiPerc, mDecompPerc, vShortNameStates, sDirectoryOutput, nYear, nStates, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []

    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = [[], []]
    vNameColsReg[0] = ['Multiplicador de produção Total', 'Multiplicador de produção Total']
    vNameColsReg[1] = ['IntraRegional', 'InterRegional']
    vDataSheet.append(mProdMultiPerc)
    vSheetName.append('ProdMultIntraInter')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = vShortNameStates[0:nStates]+["RM"]
    vDataSheet.append(mDecompPerc)
    vSheetName.append('Decomposição')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    sFileSheet = 'IntraInter_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteChock(tTotImpactI, sChockName, vShortNameStates, sDirectoryOutput, nYear, nStates, nSectors):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = ['VA_I', 'Occup_I', 'Remun_I', 'EOB_I', 'Salary_I', 'Production_I']
    xI=np.concatenate(tTotImpactI, axis=1)
    vDataSheet.append(xI)
    vSheetName.append('Impacto_I')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    sFileSheet = 'Chock_National_'+ sChockName +'_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteChockState(tTotImpactIState, sChockName, vShortNameStates, sDirectoryOutput, nYear, nStates, nSectors, nNumCoef):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vResumo = np.zeros([nStates, nNumCoef], dtype=float)
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = ['VA_I', 'Occup_I', 'Remun_I', 'EOB_I', 'Salary_I','Production_I']
    for s in range(nStates):
        mAux = tTotImpactIState[s, :, :]
        vResumo[s,:] = np.sum(mAux, axis=0)
        vDataSheet.append(mAux)
        vSheetName.append('Choque_' + vShortNameStates[s])
        vRowsLabel.append(vNameRowsReg)
        vColsLabel.append(vNameColsReg)
        vUseHeader.append(True)

    vDataSheet.append(vResumo)
    vSheetName.append('Resumo')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)


    sFileSheet = 'Chock_State_' + sChockName +'_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteChockRegion(tTotImpactIRegion, sChockName, vNameRegions, vShortNameStates, sDirectoryOutput, nYear, nStates, nSectors, nRegions, nNumCoef):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vResumo= np.zeros([nRegions, nNumCoef], dtype=float)
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = ['VA_I', 'Occup_I', 'Remun_I', 'EOB_I', 'Salary_I','Production_I']
    for s in range(nRegions):
        mAux = tTotImpactIRegion[s, :, :]
        vResumo[s,:] = np.sum(mAux, axis=0)
        vDataSheet.append(mAux)
        vSheetName.append('Choque_' + vNameRegions[s])
        vRowsLabel.append(vNameRowsReg)
        vColsLabel.append(vNameColsReg)
        vUseHeader.append(True)

    vNameRowsReg=vNameRegions
    vDataSheet.append(vResumo)
    vSheetName.append('Resumo')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    sFileSheet = 'Chock_Region_' + sChockName +'_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return

def WriteChockValues(mChock, vShortNameStates, vNameSectorReg, mChockName, sDirectoryOutput, nYear, nStates, nSectors, nChocks):
    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []


#    vNameRowsReg = [[], []]
#    vNameRowsReg[0] = vNameSectorReg[0].copy()
#    vNameRowsReg[1] = vNameSectorReg[1].copy()
    vNameRowsReg = vShortNameStates[0:nStates]
    vNameColsReg = mChockName
    mNet = np.zeros([nStates, nChocks], dtype=float)
    for c in range(nChocks):
        for s in range(nStates):
            mNet[s,c] = np.sum(mChock[s,s * nSectors:(s + 1) * nSectors, c])

    vDataSheet.append(mNet)
    vSheetName.append('ChockNet')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)
    mResumo= np.zeros([nStates, nChocks], dtype=float)
    for c in range(nChocks):
        for s in range(nStates):
            mResumo[s,c] = np.sum(mChock[:,s * nSectors:(s + 1) * nSectors, c])

    vDataSheet.append(mResumo)
    vSheetName.append('ChockNational')
    vRowsLabel.append(vNameRowsReg)
    vColsLabel.append(vNameColsReg)
    vUseHeader.append(True)

    sFileSheet = 'Chock_Values_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)
    return