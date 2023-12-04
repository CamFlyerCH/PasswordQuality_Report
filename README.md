## PasswordQuality Report

These two scripts help to create a nice report from the output of [DSInternals](https://github.com/MichaelGrafnetter/DSInternals/blob/master/Documentation/PowerShell/Readme.md) [Test-PasswordQuality](https://www.dsinternals.com/en/auditing-active-directory-password-quality/) function.

There are 2 scripts, because I think the procedure should be done in 2 steps:

1.  Data-Collection on the Domain Controller with the installed DSInternals PowerShell Module
2.  Creation of the Excel report from the collected data with the help of the excelent [ImportExcel](https://github.com/dfinke/ImportExcel) PowerShell module to import/export Excel spreadsheets, without the need of Excel itself. 

Check the Demo folder for a sample of the created files.

### Step by step guide

#### Download an updated list of leakad passwords (hashes)

To
