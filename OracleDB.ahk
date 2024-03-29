﻿#Requires AutoHotkey v2.0

Provider := "OraOLEDB.Oracle.1"
Protocol := "TCP"
Server := "DEDICATED"

; Some examples

; UpdateExample() {
;     Connection := OpenDBConnection("192.168.0.100", "1521", "MAIN", "admin", "admin")
;     ExecuteQuery(Connection, "UPDATE table SET column = 1 WHERE day = 1")
;     Connection.Close()
; }

; SelectExample() {
;     Connection := OpenDBConnection("192.168.0.100", "1521", "MAIN", "admin", "admin")
;     Recordset := ExecuteQuery(Connection, "SELECT 1 FROM DUAL")
;     MsgBox Recordset.Fields.Item(0).Value
;     Recordset.Close()
;     Connection.Close()
; }

OpenDBConnection(Host, Port, ServiceName, User, Password) {
    try {
        Connection := ComObject("ADODB.Connection")
        Connection.ConnectionString := Format("Provider={};Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL={})(HOST={})(PORT={}))(CONNECT_DATA=(SERVER={})(SERVICE_NAME={})));user ID={};password={};", Provider, Protocol, Host, Port, Server, ServiceName, User, Password)
        Connection.Open()
        return Connection
    } catch as e {
        MsgBox "Error connecting to the database: " e.Message
        ExitApp
    }
}

ExecuteQuery(Connection, SqlQuery) {
    try {
        if (SubStr(SqlQuery, 1, 6) == "SELECT") {
            Recordset := ComObject("ADODB.Recordset")
            Recordset.Open(SqlQuery, Connection)
            return Recordset
        } else {
            Connection.Execute(SqlQuery)
        }
    } catch as e {
        MsgBox "Error executing the SQL query: " e.Message
        ExitApp
    }
}
