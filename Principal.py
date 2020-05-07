# ============================================================================================
# Main Module of Building Input output Regional
# Coded to Python by João Maria de Oliveira -
# ============================================================================================
import yaml
import numpy as np
import SupportFunctions as Support
import WriteMatrix
import sys
import time

# ============================================================================================
# read Parameters into Confg file
# ============================================================================================
conf = yaml.load(open('config.yaml', 'r'), Loader=yaml.FullLoader)

sDirectoryInput  = conf['sDirectoryInput']
sDirectoryOutput = conf['sDirectoryOutput']

nColsDemand = conf['nColsDemand']
nColsDemandEach = conf['nColsDemandEach']
nColsOffer = conf['nColsOffer']
nRowsAV = conf['nRowsAV']
nColExport = conf['nColExport']
nColISFLSFConsum = conf['nColISFLSFConsum']
nColGovernConsum = conf['nColGovernConsum']
nColFamilyConsum = conf['nColFamilyConsum']
nColFBCF = conf['nColFBCF']
nColStockVar = conf['nColStockVar']
nColMarginTrade = conf['nColMarginTrade']
nColMarginTransport = conf['nColMarginTransport']
nColIPI = conf['nColIPI']
nColICMS = conf['nColICMS']
nColOtherTaxes = conf['nColOtherTaxes']
nColImport = conf['nColImport']
nColImportTax = conf['nColImportTax']

nDimension = 3
nYear = 2013
nStates = 27
nRegions = 5
nGrupSectors = 18
# Number of variables analised
nNumCoef = 6

nLinTaxes = 7
nLinRemun = 0

#vValChock = [ 0.01, 0.01, 0.01]
vValChock = [ 0.0611, 0.01, 0.0398]
mChockName = ['Export', 'Consum', 'Invest']
lAllDemand = True
lZerolittles = False
nLinWages = 0
nLinTotalProduction = 13
nLinEOBTotal = 6
nLinEOBPure = 8
nLinRMB = 7
nLinVA = 12
nLinOccup = 14

vSectors  = [12, 20,  51,  68]

# nProducts - Número de produtos de acordo com a dimensão da MIP
# nSectors - Número de atividades de acordo com a dimensão da MIP
# lAdjustMargins - True se ajusta as margens de comércio e transporte para apenas um produto e uma atividade
nSectors = vSectors [nDimension]

lWriteCoefsMatrix=False
lWriteLeontief = False
lWriteGhosh = False
lWriteMultipliers= True

# nRowTotalProduction - Número da linha do total da produção
# nRowAddedValueGross - Número da Linha do valor adicionado Bruto
# nColTotalDemand - Numero da coluna da demanda total na  tabela demanda
# nColFinalDemand - Numero da coluna da demanda total na  tabela demanda
nRowTotalProduction = nRowsAV - 2
nRowAddedValueGross = 0
nColTotalDemand = nColsDemand - 1
nColFinalDemand = nColsDemand - 2
nColsDemandReg=6

# sFileUses             - Arquivo de usos - Demanda
# sFileResources        - Arquivo de recursos
# sFileSheet - nome do arquivo de saida contendo as = tabelas
sFileInputNational ='CabecalhosMIPExtracaoRegional.xlsx'
sFileInputRegional ='MIP_Reg_JM_2013_68.xlsx'
sSheet="MIP_Reg"
sSheetNational='BR'
sFileStates= 'TabelaEstados.xlsx'
sFileShock='Choque.xlsx'
lRelativAbsolut=True

if __name__ == '__main__':

    nBeginModel = time.perf_counter()
    sTimeBeginModel = time.localtime()
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    print("Generating IO Analisys for ", nSectors, "sectors to ", nStates," States")
    print("Begin at ", time.strftime("%d/%b/%Y - %H:%M:%S",sTimeBeginModel ))
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    if lRelativAbsolut:
        nAdjustUni=1
    else:
        nAdjustUni=0

    vCodStates, vRegionState, vNameStates, vShortNameStates = Support.load_table_states(sDirectoryInput, sFileStates, nStates)
    vNameRegions = ['Norte','Nordeste','Sudeste','Sul','CentroOeste']
    vCodSector, vNameSector, vNameTaxes, vNameVA, vNameDemand, = \
        Support.load_CabecalhoMIP(sDirectoryInput, sFileInputNational, sSheetNational, nSectors)

    vCodGrupSector, vNameGrupSector =Support.load_GrupSector(sDirectoryInput, "SetoresGrupo.xlsx", nGrupSectors, nSectors)
    mDemandShock      = Support.load_DemandShock(sDirectoryInput, sFileShock, "Demanda", nSectors, nAdjustUni)
    vLaborRestriction = Support.load_OfferShock(sDirectoryInput, sFileShock, "Oferta", nSectors, nAdjustUni)

    mDemandShockReg = np.tile(mDemandShock, (nStates, nStates))

    vNameSectorReg = [[], []]
    for s in range(nStates):
        for e in range(nSectors):
            if (e == 0) or (e == nSectors - 1):
                vNameSectorReg[0].append(vShortNameStates[s])
            else:
                vNameSectorReg[0].append(vShortNameStates[s])

            #            vNameSectorReg[1].append(vCodSector[e])
            vNameSectorReg[1].append(vNameSector[e])
    # ============================================================================================
    # Import values from Regional MIP
    # ============================================================================================
    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Reading Regional Mip Matrix")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")

    mIntermConsumReg, vNatProductReg, vImportsReg, mTaxesReg, vIntermConsumTotProdReg, mVAReg, \
    vIntermConsumTotConsReg, vProdIntermConsumNatTotConsumReg, vImportsConsumNatTotReg, vTotTaxesIntermConsumReg, vIntermConsumTotConsumTotReg, vTotVAIntermConsumReg, \
    mDemandReg, vNatDemandReg, vImportsDemandReg, mTaxesDemandReg, vTotDemandReg, mVADemandReg, \
    vTotFinalDemandReg, vTotProdNacFinalDemandReg, vTotImportsFinalDemandReg, vTotTaxesFinalDemandReg, vTotVAFinalDemandReg, vTotIntermConsumFinalDemandReg, \
    vTotalDemandReg, vTotProdNacDemandReg, vTotImportsDemandReg, vTotalTaxesReg, vTotIntermConsumDemandReg, vTotalVAReg = \
        Support.load_RegionalMIP (sDirectoryInput, sFileInputRegional, sSheet,nStates,  nSectors)

    # ============================================================================================
    # Building Regional IO Matrix
    # ============================================================================================
    mInterConsumRegAux = np.concatenate((mIntermConsumReg, vNatProductReg, vImportsReg, mTaxesReg, vIntermConsumTotProdReg, mVAReg), axis=0)
    mDemandRegAux = np.concatenate((mDemandReg, vNatDemandReg, vImportsDemandReg, mTaxesDemandReg, vTotDemandReg, mVADemandReg), axis=0)
    vInterConsumTotalRegAux = np.concatenate((vIntermConsumTotConsReg, vProdIntermConsumNatTotConsumReg, vImportsConsumNatTotReg, vTotTaxesIntermConsumReg, vIntermConsumTotConsumTotReg, vTotVAIntermConsumReg), axis=0)
    vFinalDemandAuxReg = np.concatenate((vTotFinalDemandReg, vTotProdNacFinalDemandReg, vTotImportsFinalDemandReg, vTotTaxesFinalDemandReg, vTotIntermConsumFinalDemandReg, vTotVAFinalDemandReg ),axis=0)
    vTotalDemandAuxReg = np.concatenate(( vTotalDemandReg, vTotProdNacDemandReg, vTotImportsDemandReg, vTotalTaxesReg, vTotIntermConsumDemandReg, vTotalVAReg), axis=0)
    mMIPGeralReg= np.concatenate((mInterConsumRegAux, vInterConsumTotalRegAux, mDemandRegAux, vFinalDemandAuxReg, vTotalDemandAuxReg), axis=1)
    del mInterConsumRegAux, mDemandRegAux, vInterConsumTotalRegAux, vFinalDemandAuxReg, vTotalDemandAuxReg

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Calculating Regional Analisys")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    nStatesSectors = nStates * nSectors

    # Montando Matriz A e Matriz B Preparando Para Leontief e Ghosh
    vTotalProduction = mVAReg[nLinTotalProduction,:].reshape(1,nStatesSectors)
    mZ = np.copy(mIntermConsumReg)
    mA = np.zeros([nStatesSectors, nStatesSectors], dtype=float)
    for j in range(nStatesSectors):
        if (lZerolittles and vTotalProduction[0,j]<=0.01):
            mA[:, j]=0
        else:
            mA[:,j] =mZ[:,j] / vTotalProduction[0,j]

    # Calculando Matrizes inversas de Leontief, Ghosh,  Leontief Barra e Ghosh Barra
    mI = np.eye(nStatesSectors)
    mLeontief = np.linalg.inv(mI - mA)


    mAGrup = np.zeros([nGrupSectors*nStates, nGrupSectors*nStates], dtype=float)
    for il in range(nStates):
        for sl in range(nSectors):
            nLin = il*nGrupSectors + vCodGrupSector[sl]
            for ic in range(nStates):
                for sc in range(nSectors):
                    nCol = ic*nGrupSectors + vCodGrupSector[sc]
                    mAGrup[nLin, nCol] = mAGrup[nLin, nCol] + mA[il*nSectors+sl, ic*nSectors+sc]

    mIGrup = np.eye(nGrupSectors*nStates)
    mX = mIGrup - mAGrup
    mLeontiefGrup = np.linalg.inv(mIGrup - mAGrup)

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Calculating Shock")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # Calculando os Coeficientes para avaliação do choque -  VA/Occ/EOB/RMB/Wages
    vVBPReg = vTotalProduction[0, :]
    mResultsNat = np.zeros([3, 6], dtype=float)
    mResultsReg = np.zeros([nStates, 3, 6], dtype=float)
    mResultsMac = np.zeros([nRegions, 3, 6], dtype=float)
    vVAReg   = mVAReg[nLinVA,:]
    vOccReg  = mVAReg[nLinOccup,:]
    vEOBReg  = mVAReg[nLinEOBPure,:]
    vRMBReg  = mVAReg[nLinRMB,:]
    vWagesReg= mVAReg[nLinWages,:]

    mResultsNat[0, 0] =np.sum(vVBPReg)
    mResultsNat[0, 1] =np.sum(vVAReg)
    mResultsNat[0, 2] =np.sum(vEOBReg)
    mResultsNat[0, 3] =np.sum(vRMBReg)
    mResultsNat[0, 4] =np.sum(vWagesReg)
    mResultsNat[0, 5] =np.sum(vOccReg)



    vV_VAReg  = vVAReg / vVBPReg
    vV_OccReg = vOccReg / vVBPReg
    vV_EOBReg = vEOBReg / vVBPReg
    vV_RMBReg = vRMBReg / vVBPReg
    vV_WagesReg = vWagesReg / vVBPReg

    # Calculando o VBP Normal
    vFinalDemandRegNorm=np.sum(mDemandReg, axis=1)
    vVBPNormReg = np.dot(mLeontief, vFinalDemandRegNorm)

    # Aplicando o choque da demanda
    mDemandRegShock = mDemandReg * mDemandShockReg

    # Calculando o VBP do choque
    vFinalDemandRegShock=np.sum(mDemandRegShock, axis=1)
    vVBPShockDemandReg = np.dot(mLeontief, vFinalDemandRegShock)

    # Calculando o impacto do choque sobre o VBP
    vDeltaShockDemandReg = vVBPShockDemandReg - vVBPNormReg
    vImpactVBP        = vDeltaShockDemandReg / vVBPNormReg

    # CaLculando o impacto do choque
    vDeltaVAReg    = vV_VAReg * vDeltaShockDemandReg
    vDeltaEOBReg   = vV_EOBReg * vDeltaShockDemandReg
    vDeltaRMBReg   = vV_RMBReg * vDeltaShockDemandReg
    vDeltaWagesReg = vV_WagesReg * vDeltaShockDemandReg
    vDeltaOccReg   = vV_OccReg * vDeltaShockDemandReg

    mResultsNat[1, 0]   = np.sum(vDeltaShockDemandReg)
    mResultsNat[1, 1]   = np.sum(vDeltaVAReg)
    mResultsNat[1, 2]   = np.sum(vDeltaEOBReg)
    mResultsNat[1, 3]   = np.sum(vDeltaRMBReg)
    mResultsNat[1, 4]   = np.sum(vDeltaWagesReg)
    mResultsNat[1, 5]   = np.sum(vDeltaOccReg)
    mResultsNat[2,:]    = mResultsNat[1,:]  / mResultsNat[0,:] * 100
    for s in range(nStates):
        mResultsReg[s, 0, 0] = np.sum(vVBPReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 0, 1] = np.sum(vVAReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 0, 2] = np.sum(vEOBReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 0, 3] = np.sum(vRMBReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 0, 4] = np.sum(vWagesReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 0, 5] = np.sum(vOccReg[s*nSectors:(s+1)*nSectors])

        mResultsReg[s, 1, 0] = np.sum(vDeltaShockDemandReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 1, 1] = np.sum(vDeltaVAReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 1, 2] = np.sum(vDeltaEOBReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 1, 3] = np.sum(vDeltaRMBReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 1, 4] = np.sum(vDeltaWagesReg[s*nSectors:(s+1)*nSectors])
        mResultsReg[s, 1, 5] = np.sum(vDeltaOccReg[s*nSectors:(s+1)*nSectors])

        mResultsReg[s, 2, :] = mResultsReg[s, 1, :] / mResultsReg[s, 0, :] * 100

        for l in range(2):
            for c in range(6):
                mResultsMac[vRegionState[s], l, c] += mResultsReg[s, l, c]

    for r in range(nRegions):
        mResultsMac[r, 2, :] = mResultsMac[r, 1, :] / mResultsMac[r, 0, :] * 100

    vDataSheet = []
    vSheetName = []
    vRowsLabel = []
    vColsLabel = []
    vUseHeader = []
    vNameCols1 = ['Exportações', 'Governo', 'ISFLSF' ,'Familia', 'FCBF', 'Estoque' ]
    vDataSheet.append(mDemandShock)
    vSheetName.append("ChoqueDemanda")
    vRowsLabel.append(vNameSector)
    vColsLabel.append(vNameCols1)
    vUseHeader.append(True)

    vNameCols2 = ['Restrição Trabalho']
    vDataSheet.append(vLaborRestriction)
    vSheetName.append("Oferta")
    vRowsLabel.append(vNameSector)
    vColsLabel.append(vNameCols2)
    vUseHeader.append(True)

    vNameCols4 = ['VBP', 'PIB', 'EOB', 'RMB', 'Salários','Ocupações']
    vNameRows =  ['MIP', 'Impactos', 'Variação %']
    vDataSheet.append(mResultsNat)
    vSheetName.append("ImpactosNacionais")
    vRowsLabel.append(vNameRows)
    vColsLabel.append(vNameCols4)
    vUseHeader.append(True)

    mResults = mResultsReg[0, :, :]
    vNamesRowsAll=[[],[]]
    for i in range(3):
        vNamesRowsAll[0].append(vShortNameStates[0])
        vNamesRowsAll[1].append(vNameRows[i])
    for s in range(1,nStates):
        mResults = np.vstack((mResults,mResultsReg[s, :, :]))
        for i in range(3):
            vNamesRowsAll[0].append(vShortNameStates[s])
            vNamesRowsAll[1].append(vNameRows[i])


    vDataSheet.append(mResults)
    vSheetName.append("ImpactosPorUF")
    vRowsLabel.append(vNamesRowsAll)
    vColsLabel.append(vNameCols4)
    vUseHeader.append(True)

    mResultsAux = mResultsMac[0, :, :]
    vNamesRowsMac = [[], []]
    for i in range(3):
        vNamesRowsMac[0].append(vNameRegions[0])
        vNamesRowsMac[1].append(vNameRows[i])
    for c in range(1, nRegions):
        mResultsAux = np.vstack((mResultsAux, mResultsMac[c, :, :]))
        for i in range(3):
            vNamesRowsMac[0].append(vNameRegions[c])
            vNamesRowsMac[1].append(vNameRows[i])
    vDataSheet.append(mResultsAux)
    vSheetName.append("ImpactosPorREgião")
    vRowsLabel.append(vNamesRowsMac)
    vColsLabel.append(vNameCols4)
    vUseHeader.append(True)

    sFileSheet = 'Chock_Regional_' + str(nYear) + '_' + str(nSectors) + '.xlsx'
    Support.write_data_excel(sDirectoryOutput, sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel,
                             vUseHeader)

    sTimeIntermediate = time.localtime()
    print(time.strftime("%d/%b/%Y - %H:%M:%S", sTimeIntermediate), " - Writing Regional Analisys ")
    print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")



    # ============================================================================================
    # writing  Choks for state
    # ============================================================================================


    print("Terminou ")
    sys.exit(0)
'''
    # ============================================================================================
    # writing  Tech Coefs Leontiefs  and Ghoshs Matrix
    # ============================================================================================
    if lWriteCoefsMatrix:
        WriteMatrix.WriteCoefsMatrix(mA, mABarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors)
    if lWriteLeontief:
        WriteMatrix.WriteLeontief(mLeontief, mLeontiefBarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors)
    if lWriteGhosh:
        WriteMatrix.WriteGhosh(mGhosh, mGhoshBarr, vNameSectorReg, sDirectoryOutput, nYear, nSectors)
    # Calculando Multiplicadores de produção e decomposição da produção regional
    vProdMultLeontiefII = np.sum(mLeontiefBarr, axis=0).reshape(1,nStatesSectors+1)
    vProdMultGhoshII    = np.sum(mGhoshBarr, axis=0).reshape(1, nStatesSectors+1)
    vProdMultLeontiefI  = np.sum(mLeontief, axis=0).reshape(1,nStatesSectors)
    vProdMultGhoshI      = np.sum(mGhosh, axis=0).reshape(1,nStatesSectors)

    WriteMatrix.WriteMultiplierProduction(vProdMultLeontiefI, vProdMultLeontiefII, vShortNameStates, vNameSector, sDirectoryOutput,
                                          nYear, nStates, nSectors, nStatesSectors)

    # decomposição da produção regional
    mDemandFin = np.zeros([nStatesSectors, nStates+1], dtype=float)
    mTotRow = np.zeros([nStatesSectors, 1], dtype=float)
    for i in range(nStates):
        mDemandFin[:,nStates] = mDemandFin[:,nStates] + mTotEachDemand[i,:,nColExport]
        mDemandFin[:,i] = np.sum(mTotEachDemand[i,:,1:nColsDemandReg], axis=1)

    mDecomp= (np.dot(mLeontief,mDemandFin))

    mDecompPerc= np.zeros([nStates, nStates+1], dtype=float)
    for i in range(nStates):
        nTotLin = np.sum(mDecomp[i*nSectors:(i+1)*nSectors,:])
        for c in range(nStates+1):
            mDecompPerc[i,c] = np.sum(mDecomp[i*nSectors:(i+1)*nSectors,c]) / nTotLin * 100

    mProdMulti = np.zeros([nStates, 2], dtype=float)
    mProdMultiPerc =  np.zeros([nStates, 2], dtype=float)
    for i in range(nStates):
        mProdMulti[i, 0] = np.sum(mLeontief[i*nSectors:(i+1)*nSectors, i*nSectors:(i+1)*nSectors])
        mProdMulti[i, 1] = np.sum(mLeontief[:, i*nSectors:(i+1)*nSectors]) -  mProdMulti[i, 0]
        mProdMultiPerc[i, 0] = mProdMulti[i, 0] / (mProdMulti[i, 0] + mProdMulti[i, 1]) * 100
        mProdMultiPerc[i, 1] = mProdMulti[i, 1] / (mProdMulti[i, 0] + mProdMulti[i, 1]) * 100

    WriteMatrix.WriteIntraInter(mProdMultiPerc, mDecompPerc, vShortNameStates, sDirectoryOutput, nYear, nStates, nSectors)

    # Calculando Coeficentes, Geradores e multiplicadores v para VA, Ocupações, remunerações e salários
    tVAI = Support.Calc_MultiplierI(mVAReg[nLinVA,:], nStatesSectors, vTotalProduction, mLeontief)
    tVAII= Support.Calc_MultiplierII(tVAI[2], nStatesSectors, mLeontiefBarr)
    vTotalProductionOccup = np.copy(vTotalProduction) * 1000000
    tOccupI = Support.Calc_MultiplierI(mVAReg[nLinOccup,:] , nStatesSectors, vTotalProductionOccup, mLeontief)
    tOccupII= Support.Calc_MultiplierII(tOccupI[2], nStatesSectors, mLeontiefBarr)
    vAux = mVAReg[nLinWages,:] + mVAReg[nLinEOBTotal,:]
    tRemunI = Support.Calc_MultiplierI(vAux, nStatesSectors, vTotalProduction, mLeontief)
    tRemunII= Support.Calc_MultiplierII(tRemunI[2], nStatesSectors, mLeontiefBarr)

    tEOBI = Support.Calc_MultiplierI(mVAReg[nLinEOBPure,:], nStatesSectors, vTotalProduction, mLeontief)
    tEOBII= Support.Calc_MultiplierII(tEOBI[2], nStatesSectors, mLeontiefBarr)

    tSalaryI = Support.Calc_MultiplierI(mVAReg[nLinWages,:], nStatesSectors, vTotalProduction, mLeontief)
    tSalaryII= Support.Calc_MultiplierII(tSalaryI[2], nStatesSectors, mLeontiefBarr)
    # ============================================================================================
    # writing  Multipliers
    # ============================================================================================
    if lWriteMultipliers:
        tTuplesIa = tVAI[0], tVAII[0], tOccupI[0], tOccupII[0], tRemunI[0], tRemunII[0], tEOBI[0], tEOBII[0], tSalaryI[0], tSalaryII[0]
        tTuplesIb = tVAI[1], tVAII[1], tOccupI[1], tOccupII[1], tRemunI[1], tRemunII[1], tEOBI[1], tEOBII[1], tSalaryI[1], tSalaryII[1]
        tTuplesII = tVAI[2], tOccupI[2], tRemunI[2], tEOBI[2], tSalaryI[2]
        WriteMatrix.WriteMultipliers(tTuplesIa, tTuplesIb, tTuplesII, vNameSectorReg, sDirectoryOutput, nYear, nSectors)

    # Calculando Backwared e Forward Linkages
    nMeanL = np.mean(mLeontief)
    nMeanG = np.mean(mGhosh)
    vBackLinkL = np.zeros([nStatesSectors, 1], dtype=float)
    vForwLinkL = np.zeros([nStatesSectors, 1], dtype=float)
    vBackLinkG = np.zeros([nStatesSectors, 1], dtype=float)
    vForwLinkG = np.zeros([nStatesSectors, 1], dtype=float)
    for i in range(nStatesSectors):
        vBackLinkL[i, 0] = sum(mLeontief[:, i]) / nStatesSectors / nMeanL
        vForwLinkL[i, 0] = sum(mLeontief[i, :]) / nStatesSectors * nMeanL
        vBackLinkG[i, 0] = sum(mGhosh[:, i]) / nStatesSectors / nMeanG
        vForwLinkG[i, 0] = sum(mGhosh[i, :]) / nStatesSectors / nMeanG

    WriteMatrix.WriteLinkages(vBackLinkL, vForwLinkG, vBackLinkG, vForwLinkL, vShortNameStates, vNameSector, sDirectoryOutput,
                              nYear, nStates, nSectors)

    # Calculando Intensity Matrix
    mIntensity = np.zeros([nStates, nSectors, nSectors], dtype=float)
    for s in range(nStates):
        mAuxLeontief = mLeontief[s * nSectors:(s + 1) * nSectors, s * nSectors:(s + 1) * nSectors]
        #        mAuxmGhosh=mGhosh[s*nSectors:(s+1)*nSectors, s*nSectors:(s+1)*nSectors]
        nSummL = np.sum(mAuxLeontief)
        vAuxc = np.sum(mAuxLeontief, keepdims=True, axis=0)
        vAuxl = np.sum(mAuxLeontief, keepdims=True, axis=1)
        mIntensity[s, :, :] = (1 / nSummL) * np.dot(vAuxl, vAuxc)

    WriteMatrix.WriteIntensityMatrix(mIntensity, vShortNameStates, vNameSector, sDirectoryOutput, nYear, nStates, nSectors)

    # Calculando Intensity Matrix Agrupada
    mIntensity = np.zeros([nStates, nGrupSectors, nGrupSectors], dtype=float)
    for s in range(nStates):
        mAuxLeontief = mLeontiefGrup[s * nGrupSectors:(s + 1) * nGrupSectors, s * nGrupSectors:(s + 1) * nGrupSectors]
        #        mAuxmGhosh=mGhosh[s*nSectors:(s+1)*nSectors, s*nSectors:(s+1)*nSectors]

        nSummL = np.sum(mAuxLeontief)
        vAuxc = np.sum(mAuxLeontief, keepdims=True, axis=0)
        vAuxl = np.sum(mAuxLeontief, keepdims=True, axis=1)
        mIntensity[s, :, :] = (1 / nSummL) * np.dot(vAuxl, vAuxc)

    WriteMatrix.WriteIntensityMatrixGrouped(mIntensity,  vShortNameStates, vNameGrupSector, sDirectoryOutput, nYear, nStates,
                                     nGrupSectors)

'''