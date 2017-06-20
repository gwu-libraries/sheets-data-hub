# sheets-data-hub
Google App script for syncing data from multiple shared spreadsheets.

## Usage

These scripts provide functions for using the Google Sheets API to read/write changes to and from a "local" sheet (accessed by the user) and a "remote" sheet, a.k.a. the "data hub." Multiple local sheets can use these functions to receive updates from the hub and broadcast their updates, via the hub, to the other sheets.

This functionality may be useful in cases where multiple users need selective read/write access to different parts of a central data set. The scripts are intended to replicate *some* of the functionality of a relational database in the cloud, without the overhead of implementing a full-stack solution (and/or where an organizational implementation of the Google App suite prohibits using Sheets with an external web service).

Usage documentation is provided within Google Sheets, in the form of a sidebar panel that the user can toggle open and closed.

Once installed, the following functionality is present:

1. When a _local_ sheet is opened, or when the user selects **Get Latest Data** from the custom menu, the `update` function is called to fetch data from the _data hub_. Any changes to data in existing local rows under columns _not_ owned by the local sheet are broadcast to those cells. Any rows present in the data hub but not locally are appended at the end of the local sheet's data.

2. When the user has made changes and selects **Update Data Hub** from the custom menu, the `update` function is called, this time pushing data from locally owned columns back to the hub, updating the appropriate rows. 

3. If the user (per the **params** tab -- see below) has permission to add new rows, and if new rows are present, the `update` function will a) assign a new unique identifier to each and b) write them to the end of the data in the data hub, so that they are available to other users at the next update. 

4. Individual users may:
  * Add/delete additional columns not specified in the **params** for each local sheet. These columns will be unaffected by updates and will not be broadcast to other users.
  * Sort/filter the local sheet as they see fit. Sort order is not affected by the update function, which relies on the unique row identifiers to identify changes between local and remote/data hub.
  * Add additional tabs to the local sheet. These will not be affected by the update function.

5. Individual users may _not_:
  * Delete rows from the local data set. Deletes are currently not supported and will be overwritten on the next update from the data hub. 

**_Caveat emptor_** Important relational-database functionality has not been **not** implemented in these scripts, including crucial elements of ACID. Specifically, _atomicity_ and _isolation_ are not guaranteed, and concurrent transactions are not currently supported. To some degree, the dangers of the latter are mitigated by Google's native document-version control, which provides a way to step back through the changes to the data made by multiple users. The scripts do implement minimal _consistency_ checks, primarily through the constraint of a unique identifier that is automatically assigned to each new row in the data (akin to the SQL concept of a primary key). Finally, _durability_ depends on the native Google application, though the scripts do include a function useful for implementing a scheduled backup of the data.

## Installation

1. Create a new Google Sheet to be the _data hub_. This sheet should contain _all_ rows and columns in the data set. 
  * The _data hub_ may have multiple tabs if necessary. Each tab must correspond to a separate tab to local sheets linking to the data hub: in other words, SQL-like joins are not possible. Tabs on local sheets represent _views_ into the data _tables_ (_i.e._, tabs) in the data hub. 

2. Populate the data hub with the initial data to be stored. 
  * The first row on each tab should be a header row with names for each column.
  * Include a column on each tab to hold the unique row identifiers. These will be sequential integers. You can name this column anything you like, but hereafter the documentation assumes that this column has been named **data_key**. 
  * If there are initial rows in the data hub, populate the **data_key** column with sequential integer values for each row with data. See [here for an example](data_hub_init.png).

3. Create each local sheet.
  * Create a tab for each tab from the data_hub to which this local sheet needs access. 
  * On each tab, create a header row including the subset of columns from the data hub to which this local sheet needs access. _Local column names must be consistent with those on the data hub_. 
  * Include a column for the **data_key**. 
  * Create a tab called **params** with the following columns:
    + sheetName
    + remoteSheetName	
    + primaryKey	
    + localColumns	
    + foreignColumns	
    + canAddKeys
  * This tab is used to identify which columns on each tab (**sheetName**) on the local sheet broadcast updates (**localColumns**) and which receive updates (**foreignColumns**), as well the **data_key** column (**primaryKey**). The final column takes a Boolean value (**TRUE/FALSE**) asserting whether the local sheet user is allowed to add rows to the data hub.
  * The **params** tab should be populated as in [this example](params.png). Note that **localColumns** and **foreignColumns** should represent non-overlapping subsets of the columns in the data hub.
  * Once complete, the **params** can be hidden and/or locked so that other users do not see it. 
  * Create a tab called **change_log**, which will record changes to existing rows made by each update call.

4. Open the data hub spreadsheet and select **Tools** -> **Script editor...** -> **Create a new project.**
  * Give the project a name in the upper left corner of the screen. This name will be how you reference the code from the other sheets, so ideally it should be a legal Javascript variable name (no spaces, not a reserved word): _e.g._, **SerialsUpdate**.
  * Copy the code from the file **data_hub_sync.js** and paste it into the code editor, overwriting the default dummy code.
  * Get the **ID** for the data hub spreadsheet. The ID is the long alphanumeric string embedded in the URL of the spreadsheet (**not** the code editor). [See here for a complete description.](https://developers.google.com/sheets/api/guides/concepts#spreadsheet_id)
  * Copy this **ID** and paste it between the single quotes in the line of code that reads `this.foreignssKey = '';` This allows the local sheets to access the data hub.
  * Save the script. Then select **File** -> **Manage versions...** -> **Save new version** from the Script Editor menu. Now the script will be available to the local spreadsheets as a library.
  * Under **File** -> **Project properties**, copy the **Project key** (which you will need for the next step).

 5. Open each local sheet. Then open the Script Editor (**Tools** -> **Script editor...**) and give the local project a name.
  * Copy the code from the file **sync_call.js** and paste it into the Script Editor, overwriting the default dummy code.
  * Select **Resources** -> **Libraries** from the Script Editor menu. Enter the _project key_ (from step 4) into the space labeled **Add a library** and click **Add**. (If working with multiple versions of the code, make sure the correct version is selected from the menu above.)
  * Each of the three functions bound to this sheet makes a call to the library you have just enabled. If necessary, change the calls to `SerialsUpdate.update` and `SerialsUpdate.showSidebar`  to reflect your library name.
  * Save the script. 

6. Each user needs to have **Edit** permissions (through the Google Sheets sharing menu) for any local sheet they will use, _as well as_ to the data hub. (The latter is necessary for the `update` script to work properly.) Users should not make changes directly to the data hub, however.

7. 

