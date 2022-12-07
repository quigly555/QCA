# QCA - Quigly's Circle App

## Summary
This app will find and measure the radius of bright circles on dark backgrounds, and vice versa, using Circle Hough Transform. The general flow is that it will first look for capsules (bright circles on white backgrounds) since they’re easier to find. Then it will iterate through every discovered capsule, crop the image to just that region, and detect the cell body (dark circle on bright background). The rest is just math and bookkeeping to organize the resulting data and output it to an excel sheet.

## Requirements
The current version is only available for Windows at this time. If there is enough interest in a mac version I should be able to compile one.

Otherwise, there should not be any additional requirements. Everything required for the current release should be packaged in the single installer file.

## Installation Instructions:
Simply download and extract the Current Version archive, then run the QCA_Installer executable.

## Citation Instructions:
If using this script, please cite https://pubmed.ncbi.nlm.nih.gov/29364243/

## Notes:
If you find bugs or have feature suggestions let me know, I'm continually working on this program to make it better.

## Features in Development:
- Recursive folder searching. I will eventually implement a switch allowing you to choose if you want your directory to search recursively or not.
