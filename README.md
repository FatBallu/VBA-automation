# VBA-automation
VBA automation files and codes

## Data analysis for experiment
- Using MS Excel .xlsm files
- Array is the most important role in the code
- External file paths are used for automated importing and exporting of data files and figures

#### Things included
- Complex numbers calculation
- Excel as UI to control: Number of rows, column(s), steps
  - Names of .csv input files are having the same pattern and easy for import, but the exporting frequency numbers can be random due to division and rounding.  
     This is fixed by using the frequencies in the array as the output file names.  
     The processes of importing and exporting are therefore merged and cannot be seperated to avoid naming problems.
- Figures export with parameters specified in code, suitable for basic presentation but not publication
  - multiple charts in a single file can be exported with loop

## Translation from Simplified Chinese to Traditional Chinese  
- Using MS Word internal translation function  
- Not recommended if there are softwares like ConvertZ is available.  
- UTF-8 codec is used, which may cause problem in format.  
- Can be used as example for concatenation and automation in files operation. 
