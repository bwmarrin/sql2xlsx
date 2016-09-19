sql2xlsx
====
[![Go report](http://goreportcard.com/badge/bwmarrin/sql2xlsx)](http://goreportcard.com/report/bwmarrin/sql2xlsx) [![Discord Gophers](https://img.shields.io/badge/Discord%20Gophers-%23info-blue.svg)](https://discord.gg/0f1SbxBZjYq9jLBk)
====
A simple program to convert SQL rows into Microsoft Excel XLSX files.

This example is built to work with MS SQL driver however it can easily be 
modified to function with any other Go SQL driver.

**For help with this program or general Go discussion, please join the [Discord 
Gophers](https://discord.gg/0f1SbxBZjYq9jLBk) chat server.**

## Install & Build

This assumes you already have a working Go environment, if not please see
[this page](https://golang.org/doc/install) first.

```sh
go get https://github.com/bwmarrin/sql2xlsx.git
cd sql2xlsx
go build
```

## Usage

All options except for are required.

```
Usage of ./sql2xlsx:
  -h string
        SQL Server hostname or IP
  -u string
        User ID
  -p string
        Password
  -s string
        SQL Query filename
  -o string
        Output filename
```

