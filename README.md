![Py-MENT-SSWD](https://github.com/Gysco/SSWD/blob/master/rsrc/img/pyment_splashart.png?raw=true)

# Python Methods applied to Environmental [Nuclear] Toxicity


[![GitHub issues](https://img.shields.io/github/issues/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/issues) [![GitHub forks](https://img.shields.io/github/forks/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/network) [![GitHub stars](https://img.shields.io/github/stars/Gysco/SSWD.svg?style=flat-square)](https://github.com/Gysco/SSWD/stargazers) [![Github All Releases](https://img.shields.io/github/downloads/Gysco/SSWD/total.svg?style=flat-square)](https://github.com/Gysco/SSWD/releases)

[![GitHub license](https://img.shields.io/badge/license-AGPL-blue.svg?style=flat-square)](https://raw.githubusercontent.com/Gysco/SSWD/master/LICENSE)

[![Travis](https://img.shields.io/travis/Gysco/SSWD.svg?style=flat-square)](https://travis-ci.org/Gysco/SSWD)
[![CircleCI](https://img.shields.io/circleci/project/github/Gysco/SSWD.svg?style=flat-square)](https://circleci.com/gh/Gysco/SSWD)

## Initial Objectives

The SSWD code aims to build species sensitivity weighted distributions (SSWD) and to calculate hazardous concentration (HC, with its confidence limits) for different reference levels (HCx, where x is the accepted fraction of affected species). The study leading to the SSWD development ([Duboudin et al, 2004](https://github.com/Gysco/SSWD/blob/master/docs/Duboudin_et_al-2004-Environmental_Toxicology_and_Chemistry.pdf)) demonstrated that the value of the HC5, the usual wanted threshold, is directly impacted by both the weight of each taxonomic group (or trophic level) and species and the statistical method used to construct the distribution. The SSWD macro was developed to allow weighting of ecotoxicity concentration data to account for redundant data for each species (or genus) and for the disproportion in the data number between the taxonomic groups (or trophic levels).

## Extended objectives

Developed to define ecotoxicological benchmarks, based on toxicity data either obtained in field investigation or in laboratory, the SSWD macro may have other useful applications. In fact the code basically provides a cumulative distribution for any kind of input data, and more interesting with the associated confidence interval. One of the goals of the macro transcription is then to make easier its use for other applications than ecotoxicity studies.


## Documentations

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
