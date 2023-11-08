# Powershell-Reports
This repository houses various reports built with PowerShell.

# Description

## Accounts Never Logged on
This report sends an email with all of the AD accounts that have never been logged on to.

## Active Directory General Information Report
Generates an HTML file with information about:

- Forest Name
- Forest Mode
- Forest Domains
- Root Domain
- Domain Naming Master
- Schema Master
- Domain SPN Suffixes
- Domain UPN Suffixes
- Global Catalog Servers
- Forest Domain Sites
- Domain Information
- Domain Net Bios Name
- Domain Mode
- Parent Domain
- Domain PDC Emulator
- Domain RID Master
- Domain Infrastructure Master
- Child Domains
- Replicated Servers

## AD Token Size Report
This report generates a customizable report of AD account token sizes.

This can take a considerable time to run depending on how many users you have but has customizable options to control the output.

## Expiring Accounts

This report loops through all users in Active Directory and calculates how close the accounts passowrd is to expiring. The default is that accounts expire every 365 days. You can tweak that value in the script.

You can configure the HTML to change whatever the email that gets sent to the users looks like.

## Windows Defender Status Report

This report will get the status of all Windows Defender installations against servers in the environment.

Requires you to configure your OU paths in the Get-RegionalOSservers function.