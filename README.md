# Using-spreadsheet-documents-as-requirements-specifications-for-automatic-software-generation
We propose an approach to generate applications from spreadsheets. This includes that the spreadsheet formulas and data tables are available in the generated application and can be worked with. The idea is to be able to use spreadsheet documents as requirements specifications for automatic software generation. There are existing approaches that do the same. However, such approaches depend on the spreadsheet which causes bad performance and limited concurrent access options. Further, there are low-code platforms available which allow to generate applications based on spreadsheets but fail when it comes to making use of existing formulas within the spreadsheet. We elaborated a new approach that is fast and can work independently from the spreadsheet after an initial import. Also, it allows to make use of formulas already defined in the spreadsheet. We call this new approach the "headless spreadsheet approach", as it uses the concept of headless spreadsheets which provide spreadsheet functionalities without GUI as in traditional spreadsheet
software. The approach consists of two stages. At the first stage, a spreadsheet is imported to an application. The data from the spreadsheet is then transferred to an attached relational database. At the second stage, the user of the application is presented with the formulas as defined in the imported spreadsheet. The formulas can be evaluated upon request with the help of a headless spreadsheet instance which is
created based on the data from the database. Correctness tests have shown that our solution is robust. Moreover, performance tests have shown that our approach is significantly faster than existing solutions which interact with a traditional spreadsheet. Also, the tests have shown that the approach is suited to be used by multiple users concurrently.

# Setup in Visual Studio
* Visual Studio 2022 is required.
* Microsoft SQL Server needs to be installed. Alternatevely, SQL Server Express LocalDB can be used for testing purposes.
* To scaffold a new database migration the following command needs to be executed within the Package Manager Console of Visual Studio: `Add-Migration InitialCreate`
* The Migration can then be applied with the following command in the Package Manager Console: `Update-Database`

# Usage
Excel documents can be uploaded to the web application. The document needs to have an additional sheet "Formulas" containing the Excel formulas which need to be available in the application.
* The documentation is outlined in Documentation/MT_Martin_Frick.pdf
* A template for the Excel document is given in Templates/spreadsheet_Template.xlsx 

# Architecture
The web application follows the MVC architecture. Following this pattern, the spreadsheet document serves as Model from which a data model is generated as the internal representation of the Model. The Controller handles user requests and updates the Model if needed. Also, the Controller passes the model to the Views to render it to the user.
 
 # To Be Done
 * It is currently not possible to consolidate several spreadsheets within the same application. For, use cases where this would be required, the approach needs to be adapted accordingly.
 * Only Microsoft Excel is currently supported as spreadsheet software as EPPlus is used to read the spreadsheet. EPPLus only allows to work with Excel. To allow it to work with other spreadsheet software as well, the implementation must be adapted.

# Sources

Microsoft SQL Server:
https://www.microsoft.com/en-us/sql-server/sql-server-downloads

SQl Server Express LocalDB:
https://learn.microsoft.com/en-us/sql/database-engine/configure-windows/sql-server-express-localdb?view=sql-server-ver16

Database Migrations Overview:
https://learn.microsoft.com/en-us/ef/core/managing-schemas/migrations/?tabs=vs

Based on HyperFormula:
https://github.com/handsontable/hyperformula

Spreadsheet service with EPPLUS:
https://github.com/EPPlusSoftware/EPPlus
