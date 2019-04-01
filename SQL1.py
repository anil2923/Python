
#INSERT INTO ATRate VALUES (7,3,'0.0','1.000000','1.000000','1.000000','1.000000')
#NSERT INTO ATRate VALUES (GroupId,RateType,NACEGroupRate,NaceClassRate,ZoneRate,WageRate,RiskLevelRate)
#GroupId	RateType	NACEGroup	NACEClass	Region	Salary	RiskCategory

import pandas as pd

df1 = pd.read_excel('ATRates.xlsx', sheet_name='ATRateTypes')

with open("AT_RATE_UPD.sql", "w+") as file:
    for i in df1.index:
        df2 = pd.read_excel('ATRates.xlsx', sheet_name=df1['RateName'][i])
        R=df1['RateType'][i]
        file.write('------------------->> INSET SCRIPT STARTED FOR : {0}<<-------------------\n'.format(df1['RateName'][i]))
        for j in df2.index:
            file.write('INSERT INTO ATRate VALUES ({0},{6},\'{1}\',\'{2}\',\'{3}\',\'{4}\',\'{5}\');\n'
                        .format(df2['GroupId'][j],  
                                df2['NACEGroup'][j], 
                                df2['NACEClass'][j], 
                                df2['Region'][j],  
                                df2['Salary'][j], 
                                df2['RiskCategory'][j],R))
        file.write('------------------->>INSERT SCRIPT ENDED FOR : {0}<<-------------------\n\n'.format(df1['RateName'][i]))

        