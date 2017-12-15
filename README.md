# PatternsAndPractices Class Libraries:

### KeyConfig\ConfigurationManagement

### Microsoft.ApplicationBlocks.ConfigurationManagement

### Microsoft.ApplicationBlocks.Data

### Microsoft.ApplicationBlocks.ExceptionManagement

### Microsoft.ApplicationBlocks.UIProcess



This is the tried-and-true "Patterns and Practices" VB.Net wrapper for SQL. 

Although old, it still remains a useful wrapper for doing those repeated operations with System.Data.SQLClient

Also included are the configuration management, exception management, and UI Process libraries.

Favorite SQL operations include:

#### Execute a SQL command with no results expected:

  ```SqlHelper.ExecuteNonQuery(myConnectionString, CommandType.StoredProcedure, strSQL, params)```

#### Run a SQL command and put the results in a DataSet (or any number of DataTables in a DataSet):

  ```Dim ds As DataSet = SqlHelper.ExecuteDataset(myConnectionString, CommandType.StoredProcedure, strSQL, params)```

#### Run a SQL command that returns a single integer value

  ```UserCt = CInt(SqlHelper.ExecuteScalar(myConnectionString, CommandType.Text, "SELECT count(*) FROM mytable"))```

Can be compiled as VB.Net, but then used by other languages such as C# (see included C# WebDemo app)


Code was was freeware from Microsoft, but apparently no longer supported / available.

'==============================================================================='
 Microsoft Configuration Management Application Block for .NET
 http://msdn.microsoft.com/library/en-us/dnbda/html/cmab.asp

 AssemblyInfo.vb

 This file contains the the definitions of assembly level attributes.

 For more information see the Configuration Management Application Block Implementation Overview. 
 
'==============================================================================='
 Copyright (C) 2000-2001 Microsoft Corporation
 All rights reserved.
 THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
 OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
 LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
 FITNESS FOR A PARTICULAR PURPOSE.
'=============================================================================='

'==============================================================================='
 Microsoft User Interface Process Application Block for .NET
 http://msdn.microsoft.com/library/en-us/dnbda/html/uip.asp

 AssemblyInfo.vb

 For more information see the User Interface Process Application Block Implementation Overview. 
 
'==============================================================================='
 Copyright (C) 2000-2001 Microsoft Corporation
 All rights reserved.
 THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
 OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
 LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR
 FITNESS FOR A PARTICULAR PURPOSE.
'=============================================================================='
