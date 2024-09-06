#Requires AutoHotkey v2.0

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
    ; Connection := OpenDBConnection("192.168.0.100", "1521", "MAIN", "admin", "admin")
    ; Recordset := ExecuteQuery(Connection, "SELECT * FROM table")
    ; if IsObject(Recordset) {
    ;     if (Recordset.EOF = false) {
    ;         MsgBox Recordset.Fields.Item(0).Value
    ;     }

    ;     Recordset.Close()
    ; }
    ; Connection.Close()
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
        Command := ComObject("ADODB.Command")
        Command.ActiveConnection := Connection
        Command.CommandText := SqlQuery

        Recordset := Command.Execute()

        if IsObject(Recordset) and Recordset.State = 1 {
            return Recordset
        } else {
            return true
        }
    } catch as e {
        MsgBox "Error executing the SQL query: " e.Message
        return
    }
}
