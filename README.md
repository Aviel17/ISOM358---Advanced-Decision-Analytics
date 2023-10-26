# Advanced Decision Analytics - Final Project

## Context
Volunteer Medical Interpretation Services (VMIS) is an Emory organization training Spanish and Portuguese-speaking students to provide interpreting services for clinics in Atlanta, bridging the linguistic barrier between patients and physicians.

Given 30 volunteers, 40 operating hours per week, and 1-hour time slots, we strived to address two central paint points:
  1. Time-consuming process to satisfy everyone's availability
  2. Creating a new schedule every semester

## Goal
Alleviate such pain points with a user-friendly Excel Add-In tool that automatically creates and optimizes the on-call schedule based on students’ preferences.

  **QR Code from Microsoft Forms --> Excel Add-In Tool --> Incoming Outlook Email**

# Technologies
* Excel
* VBA
* Microsoft Forms

## Setup
Before you use this tool, make sure to download and save the following files to your local machine:
  1. `VMIS_Scheduling.xlsx` (file with database of 30 volunteer information, but feel free to generate and download your own 30 volunteer data file)
  2. `FINAL_Model.xlsm` (main file)

## Procedure
To run the program, perform the following:
  1. Open the `FINAL_Model.xlsm` macro-enabled file
  2. Enable macros (if not already enabled)
  3. Enable Solver _and_ OpenSolver (need to install OpenSolver Linear version [here](https://opensolver.org/installing-opensolver/))
  4. Follow the prompts on the buttons while going through the program
  5. Done!
