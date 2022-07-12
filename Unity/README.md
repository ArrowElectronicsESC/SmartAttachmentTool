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

# Send Email

## Run script:
```
py GetSendEmail.py Parameter1 Parameter2 Parameter3
```
Parameter1 = csv part suggestion

Parameter2 = FAE/FSR Name

Parameter3 = year

Parameter4 = Qx

e.g.:
```
py GetSendEmail.py ".\partsuggestion.csv" "Jebasingam, Hentry" "2022" "Q3"
```
## Output:

email (e.g.: sendemail.png)

## Input:

csv file

e.g.:
```
partsuggestion.csv
```