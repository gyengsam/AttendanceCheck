object DM: TDM
  OldCreateOrder = False
  Height = 184
  Width = 431
  object FDQuery1: TFDQuery
    Connection = FDConnection1
    Left = 111
    Top = 8
  end
  object FDQuery2: TFDQuery
    Connection = FDConnection1
    Left = 191
    Top = 8
  end
  object FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink
    DriverID = 'SQLite'
    VendorHome = '.\Sqlite3'
    VendorLib = '.\Sqlite3\sqlite3.dll'
    Left = 344
    Top = 8
  end
  object FDConnection1: TFDConnection
    FetchOptions.AssignedValues = [evRowsetSize]
    FetchOptions.RowsetSize = 100000
    AfterConnect = FDConnection1AfterConnect
    AfterDisconnect = FDConnection1AfterDisconnect
    Left = 24
    Top = 8
  end
  object FDConMysql: TFDConnection
    Params.Strings = (
      'Database=church'
      'User_Name=christian'
      'Password=Test52905290!'
      'Server=localhost'
      'CharacterSet=utf8'
      'DriverID=MySQL')
    FetchOptions.AssignedValues = [evRecordCountMode]
    FetchOptions.RecordCountMode = cmTotal
    ResourceOptions.AssignedValues = [rvAutoReconnect]
    ResourceOptions.AutoReconnect = True
    LoginPrompt = False
    AfterConnect = FDConMysqlAfterConnect
    AfterDisconnect = FDConMysqlAfterDisconnect
    Left = 23
    Top = 104
  end
  object FDMysqlQuery1: TFDQuery
    Connection = FDConMysql
    Left = 95
    Top = 104
  end
  object FDMysqlQuery2: TFDQuery
    Connection = FDConMysql
    Left = 175
    Top = 104
  end
  object FDQuery3: TFDQuery
    Connection = FDConnection1
    Left = 255
    Top = 8
  end
  object FDMysqlQuery3: TFDQuery
    Connection = FDConMysql
    Left = 247
    Top = 104
  end
end
