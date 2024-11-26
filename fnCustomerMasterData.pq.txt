let
  fnCustomerMasterData = (Paste_Reports_Folder_Path_Here as text, Reports_Name_Common_Pattern_eg.CustomerMaster as text) =>
    let
      Source = Folder.Files(Paste_Reports_Folder_Path_Here), 
      FilteredFiles = Table.SelectRows(Source, each Text.StartsWith([Name], Reports_Name_Common_Pattern_eg.CustomerMaster)), 
      FilterTempolaryFiles = Table.SelectRows(FilteredFiles, each not Text.StartsWith([Name], "~$")), 
      GetExcelContent = Table.TransformColumns(FilterTempolaryFiles, {"Content", each Excel.Workbook(_)}), 
      ExpandData = Table.SelectRows(Table.ExpandTableColumn(GetExcelContent, "Content", {"Data", "Kind"}), each [Kind] = "Sheet"), 
      Headers = Table.AddColumn(ExpandData, "Custom", each Table.PromoteHeaders(Table.Skip([Data], 3))), 
      DeleteOtherCols = Table.SelectColumns(Headers, {"Custom"}), 
      MainCols = List.Select(
        Table.ColumnNames(Table.Combine(DeleteOtherCols[Custom])), 
        each _
          = "Distributor Code"
          or _ = "DMS Customer Code"
          or _ = "Customer Name"
          or _ = "Price Group"
          or _ = "global Channel"
          or _ = "Territory"
          or _ = "Route"
          or _ = "Cust Status"
          or _ = "Sub-Segment 3"
      ), 
      DataTable = Table.ExpandTableColumn(DeleteOtherCols, "Custom", MainCols), 
      SplitPriceGroup = Table.SplitColumn(
        DataTable, 
        "Price Group", 
        Splitter.SplitTextByEachDelimiter({"- "}), 
        {"Price", "Price Group"}
      ), 
      SplitChannel = Table.SplitColumn(
        SplitPriceGroup, 
        "global Channel", 
        Splitter.SplitTextByEachDelimiter({"- "}), 
        {"Channel.1", "Channel"}
      ), 
      SplitTerritory = Table.SplitColumn(
        SplitChannel, 
        "Territory", 
        Splitter.SplitTextByEachDelimiter({"- "}), 
        {"Territory.1", "Territory"}
      ), 
      SplitRoute = Table.SplitColumn(SplitTerritory, "Route", Splitter.SplitTextByEachDelimiter({"- "}), {"Route.1", "Route"}), 
      SplitSegment = Table.SplitColumn(
        SplitRoute, 
        "Sub-Segment 3", 
        Splitter.SplitTextByEachDelimiter({"- "}), 
        {"SUB SEGMENT3.1", "Segment"}
      ), 
      FilterKE = Table.SelectRows(SplitSegment, each Text.StartsWith([DMS Customer Code], "KE") and [Cust Status] = "ACTIVE"), 
      DeleteColumns = Table.RemoveColumns(
        FilterKE, 
        {"Channel.1", "Territory.1", "Route.1", "SUB SEGMENT3.1", "Price"}
      ), 
      DataType = Table.TransformColumnTypes(
        DeleteColumns, 
        {
          {"Distributor Code", type text},
          {"DMS Customer Code", type text}, 
          {"Customer Name", type text}, 
          {"Cust Status", type text}, 
          {"Price Group", type text}, 
          {"Channel", type text}, 
          {"Segment", type text}, 
          {"Territory", type text}, 
          {"Route", type text}
        }
      )
    in
      DataType
in
  fnCustomerMasterData
