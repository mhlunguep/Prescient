# Prescient Assessment

## Introduction

This readme provides an overview of the completed project for the assessments provided. The project aims to create a .NET C# application that downloads Excel files from the JSE website, processes the downloaded files, and saves the content to a SQL database. Additionally, it includes the creation of a stored procedure in MS SQL Server for generating a report based on input dates.

## Project Overview

### Assessment 1

#### Objective

The objective of Assessment 1 was to develop a .NET C# application to download Excel (.xls) files from the JSE website for the year 2023 and save the file content to a SQL database.

### Implementation

- The application was developed using .NET 8 and C#.
- Microsoft Interop for Excel was utilized to handle Excel file processing.
- A simple console application was created for initiating the download and processing tasks.
- Duplicate file downloads and database entries were avoided.
- The database schema and example data for the DailyMTM table were provided and used for data storage.

## Assessment 2

### Objective

The objective of Assessment 2 was to create a MS SQL Stored Procedure ([SP_Total_Contracts_Traded_Report]) that accepts Date-From and Date-To inputs and returns a table with specific headers.

#### Implementation

- A stored procedure named [SP_Total_Contracts_Traded_Report] was developed in MS SQL Server.
- The stored procedure accepts Date-From and Date-To inputs and generates a table with headers [File Date], [Contract], [Contracts Traded], [% Of Total Contracts Traded].
- Only contracts with a [Contracts Traded] value greater than zero are included in the report.

## Instructions for Running the Project

##### 1. Setting Up the Environment:

- Ensure you have .NET 8 installed on your system.
- Install Microsoft Office (or Excel) to enable Microsoft Interop for Excel functionality.
- Set up MS SQL Server 2010 or later for database operations.

##### 2. Cloning the Repository:

- Clone the repository to your local machine using Git or download the project as a ZIP file.

##### 3. Database Setup:

- Execute the provided SQL script to create the necessary DailyMTM table in your SQL Server database.
- The script 'Create & Data Script for Table [DailyMTM]' is on the root of this project

##### 4. Connection String

- On the repo find the appsettings.json file and change the connection string to match your database

##### 5. Running the Application:

- Open the project in your preferred IDE (e.g., Visual Studio).
- Build the solution to ensure all dependencies are resolved.
- Run the console application to initiate the download and processing tasks.

##### 6. Executing the Stored Procedure:

- Use SQL Server Management Studio or any SQL client to execute the stored procedure [SP_Total_Contracts_Traded_Report].
- This stored procedure is on the root of the repo
- Provide Date-From and Date-To inputs as parameters to generate the report.

##### 7. Testing and Validation:

- Validate the application functionality by reviewing downloaded files (check the folder named Daily_MTM_Reports that will be created after running the project), database entries, and generated reports.

### Feedback and Improvements

Your feedback on the project is highly appreciated. Please review the codebase and provide any suggestions, improvements, or issues you encounter. Feel free to open GitHub issues or reach out via email with your feedback.

##### Author: Phumlani Mhlungu

##### Date: 06 March 2024
