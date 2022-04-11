# PSWriteOffice - PowerShell Module

<p align="center">
  <a href="https://dev.azure.com/evotecpl/PSWriteOffice/_build/results?buildId=latest"><img src="https://img.shields.io/azure-devops/build/evotecpl/39c74615-8f34-4af0-a835-68dc33f9214f/14?label=Azure%20Pipelines&style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/v/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/vpre/PSWriteOffice.svg?label=powershell%20gallery%20preview&colorB=yellow&style=flat-square"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/license/EvotecIT/PSWriteOffice.svg?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/p/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/languages/top/evotecit/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/languages/code-size/evotecit/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/dt/PSWriteOffice.svg?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=flat-square&logo=twitter"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg?style=flat-square"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn&style=flat-square"></a>
</p>

`PSWriteOffice` is cross-platform PowerShell module for writing Microsoft Office documents (`Word`, `Excel` and later on `PowerPoint`).
This module long term goal is to fully replace:

- [PSWriteWord](https://github.com/EvotecIT/PSWriteWord) - this module works only in Windows PowerShell 5.1, and DLL license has changed making updates to it very hard.
- [PSWriteExcel](https://github.com/EvotecIT/PSWriteExcel) - this module worked cross-platform but DLL maintainer moved to pay project leaving the old DLL unmaintained.

While at this moment this module doesn't expose all DLL features, and some features are strictly available in by using their net methods, the goal is to expose it fully using PowerShell.