# icc3 Part and Property

## Dependencies:

* pandas:
```
pip install pandas
```
  
* xlsxwriter:
```
pip install XlsxWriter
```
## Run script:
```
py icc3.py Parameter1 Parameter2
```
Parameter1 = Main Part

Parameter2 = Property

e.g.:
```
py icc3.py Bluetooth Antenna
```
## Output:

excel file: 

e.g.:

`BluetoothAntenna20220703125230.xlsx`

## Input:

excel file:

`UnityReport.xlsx`

# Get FSR and FAE Email

## Run script:
```
py getArrowEmail.py Parameter1
```
Parameter1 = FAE (faeName column) or FSR (fsrName column) Name from `BluetoothAntenna20220703125230.xlsx` excel file

e.g.:
```
py getArrowEmail.py "Jebasingam, Hentry"
```
## Output:

email: 

e.g.:

`Hentry.Jebasingam@arrow.com`
