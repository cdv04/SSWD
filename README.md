![Py-MENT-SSWD](https://github.com/Gysco/SSWD/blob/master/rsrc/img/pyment_splashart.png?raw=true)

# Python Methods applied to Environmental [Nuclear] Toxicity


[![GitHub issues](https://img.shields.io/github/issues/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/issues) [![GitHub forks](https://img.shields.io/github/forks/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/network) [![GitHub stars](https://img.shields.io/github/stars/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/stargazers) [![Github All Releases](https://img.shields.io/github/downloads/Gysco/SSWD/total.svg?style=flat-square)](https://github.com/Gysco/SSWD/releases)

[![GitHub license](https://img.shields.io/badge/license-AGPL-blue.svg?style=flat-square)](https://raw.githubusercontent.com/Gysco/SSWD/master/LICENSE)

[![Travis](https://img.shields.io/travis/Gysco/SSWD.svg?style=flat-square)](https://travis-ci.org/Gysco/SSWD)
[![AppVeyor](https://img.shields.io/appveyor/ci/Gysco/sswd.svg?style=flat-square)](https://ci.appveyor.com/project/Gysco/sswd/)

## Context

Determining protection thresholds for wildlife is a required step in the framework of ecological risk assessment. Afferent methods are proposed wordlwide, among which the so-called Species Sensitivity Distribution makes consensus as soon as the quantity and quality of ecotoxicity data are sufficient (thay may be interpreted differently depending on the context).
In 2003, the SSWD Excel Macro was developed by Electricité de France (EDF) in collaboration with Institut de l’Environnement Industriel et des Risques (INERIS). This useful tool aimed to build Species Sensitivity Weighted Distributions (SSWD) and to calculate Hazardous Concentration (HC, with its 90% confidence interval) for different reference levels (HCx, where x is the accepted fraction of affected species). However, this macro wasn't updated nor maintained since its inital release. It is now not more compatible with current versions of Windows and Excel. Py-ME[N]T-SSWD is the IRSN solution to this problem. This stand-alone systemless application, which upgrades the former Excel macro, is the first piece of a series of methodological developments in the field of environmental toxicity, gathered in the plateform Py-ME[N]T (Python Methods applied to Environmental [Nuclear] Toxicity). Developed to meet the Institute needs related to Environmental Nuclear Toxicology, they are inspired from (and then applicable to) "conventional" ecotoxicology, hence the [N]...

## Initial Objectives

The primary goal of Py-ME[N]T-SSWD is to reproduce the calculation done by the Excel macro based on the paper from [Duboudin et al, 2004](https://github.com/Gysco/SSWD/blob/master/docs/Duboudin_et_al-2004-Environmental_Toxicology_and_Chemistry.pdf)). This study demonstrated that the value of the HC5, the usual wanted threshold, is directly impacted by both the weight of each taxonomic group (or trophic level) and species and the statistical method used to construct the distribution. The SSWD macro was developed to allow weighting of ecotoxicity concentration data to account for redundant data for each species (or genus) and for the disproportion in the data number between the taxonomic groups (or trophic levels). Py-ME[N]T-SSWD runs the same algorythm, with improvements regarding the limitation in terms of data number (no such limitation in Py-ME[N]T-SSWD)...  

## Extended objectives

Developed to define ecotoxicological benchmarks, based on toxicity data either obtained in field investigation or in laboratory, Py-ME[N]T-SSWD may have other useful applications. In fact the code basically provides a cumulative distribution for any kind of input data, and more interesting with the associated confidence interval. One of the secondary goals of the macro transcription is then to make easier the use of Py-ME[N]T-SSWD for other applications than ecotoxicity studies.


## Documentation

If you want to read about using PyME[N]T-SSWD or contributing to the development, the [PyMENT-SSWD wiki](https://github.com/Gysco/SSWD/wiki) is free and available online.

## Installing

### Prerequisites

#### Building from sources
- Python 3

Install dependencies for your version of python by running:
```bash
pip install -r requirements.txt
```
:warning: **Windows user you first need to install [numpy+mkl](http://www.lfd.uci.edu/~gohlke/pythonlibs/#numpy) and [scipy](http://www.lfd.uci.edu/~gohlke/pythonlibs/#scipy)**

#### Using binaries
They are standalone version, you only need to [download](https://github.com/Gysco/SSWD/releases) them and run them.
