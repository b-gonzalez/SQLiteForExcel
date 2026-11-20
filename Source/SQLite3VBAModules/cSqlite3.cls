VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSqlite3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Notes:
' Microsoft uses UTF-16, little endian byte order.

Implements ISqlite3

Private Const P_MASTER_DATABASE_DIRECTORY As String = ""

Private Const P_MASTER_DIRECTORY_SET As Long = vbObjectError + 513

Private pDatabaseDirectory As String

Private Const JULIANDAY_OFFSET As Double = 2415018.5

' Returned from SQLite3Initialize
Private Const P_SQLITE_INIT_OK     As Long = 0
Private Const P_SQLITE_INIT_ERROR  As Long = 1

' SQLite data types
Private Const P_SQLITE_INTEGER  As Long = 1
Private Const P_SQLITE_FLOAT    As Long = 2
Private Const P_SQLITE_TEXT     As Long = 3
Private Const P_SQLITE_BLOB     As Long = 4
Private Const P_SQLITE_NULL     As Long = 5

' SQLite atandard return value
Private Const P_SQLITE_OK          As Long = 0   ' Successful result
Private Const P_SQLITE_ERROR       As Long = 1   ' SQL error or missing database
Private Const P_SQLITE_INTERNAL    As Long = 2   ' Internal logic error in SQLite
Private Const P_SQLITE_PERM        As Long = 3   ' Access permission denied
Private Const P_SQLITE_ABORT       As Long = 4   ' Callback routine requested an abort
Private Const P_SQLITE_BUSY        As Long = 5   ' The database file is locked
Private Const P_SQLITE_LOCKED      As Long = 6   ' A table in the database is locked
Private Const P_SQLITE_NOMEM       As Long = 7   ' A malloc() failed
Private Const P_SQLITE_READONLY    As Long = 8   ' Attempt to write a readonly database
Private Const P_SQLITE_INTERRUPT   As Long = 9   ' Operation terminated by sqlite3_interrupt()
Private Const P_SQLITE_IOERR      As Long = 10   ' Some kind of disk I/O error occurred
Private Const P_SQLITE_CORRUPT    As Long = 11   ' The database disk image is malformed
Private Const P_SQLITE_NOTFOUND   As Long = 12   ' NOT USED. Table or record not found
Private Const P_SQLITE_FULL       As Long = 13   ' Insertion failed because database is full
Private Const P_SQLITE_CANTOPEN   As Long = 14   ' Unable to open the database file
Private Const P_SQLITE_PROTOCOL   As Long = 15   ' NOT USED. Database lock protocol error
Private Const P_SQLITE_EMPTY      As Long = 16   ' Database is empty
Private Const P_SQLITE_SCHEMA     As Long = 17   ' The database schema changed
Private Const P_SQLITE_TOOBIG     As Long = 18   ' String or BLOB exceeds size limit
Private Const P_SQLITE_CONSTRAINT As Long = 19   ' Abort due to constraint violation
Private Const P_SQLITE_MISMATCH   As Long = 20   ' Data type mismatch
Private Const P_SQLITE_MISUSE     As Long = 21   ' Library used incorrectly
Private Const P_SQLITE_NOLFS      As Long = 22   ' Uses OS features not supported on host
Private Const P_SQLITE_AUTH       As Long = 23   ' Authorization denied
Private Const P_SQLITE_FORMAT     As Long = 24   ' Auxiliary database format error
Private Const P_SQLITE_RANGE      As Long = 25   ' 2nd parameter to sqlite3_bind out of range
Private Const P_SQLITE_NOTADB     As Long = 26   ' File opened that is not a database file
Private Const P_SQLITE_ROW        As Long = 100  ' sqlite3_step() has another row ready
Private Const P_SQLITE_DONE       As Long = 101  ' sqlite3_step() has finished executing

' Extended error codes
Private Const P_SQLITE_IOERR_READ               As Long = 266  '(SQLITE_IOERR | (1<<8))
Private Const P_SQLITE_IOERR_SHORT_READ         As Long = 522  '(SQLITE_IOERR | (2<<8))
Private Const P_SQLITE_IOERR_WRITE              As Long = 778  '(SQLITE_IOERR | (3<<8))
Private Const P_SQLITE_IOERR_FSYNC              As Long = 1034 '(SQLITE_IOERR | (4<<8))
Private Const P_SQLITE_IOERR_DIR_FSYNC          As Long = 1290 '(SQLITE_IOERR | (5<<8))
Private Const P_SQLITE_IOERR_TRUNCATE           As Long = 1546 '(SQLITE_IOERR | (6<<8))
Private Const P_SQLITE_IOERR_FSTAT              As Long = 1802 '(SQLITE_IOERR | (7<<8))
Private Const P_SQLITE_IOERR_UNLOCK             As Long = 2058 '(SQLITE_IOERR | (8<<8))
Private Const P_SQLITE_IOERR_RDLOCK             As Long = 2314 '(SQLITE_IOERR | (9<<8))
Private Const P_SQLITE_IOERR_DELETE             As Long = 2570 '(SQLITE_IOERR | (10<<8))
Private Const P_SQLITE_IOERR_BLOCKED            As Long = 2826 '(SQLITE_IOERR | (11<<8))
Private Const P_SQLITE_IOERR_NOMEM              As Long = 3082 '(SQLITE_IOERR | (12<<8))
Private Const P_SQLITE_IOERR_ACCESS             As Long = 3338 '(SQLITE_IOERR | (13<<8))
Private Const P_SQLITE_IOERR_CHECKRESERVEDLOCK  As Long = 3594 '(SQLITE_IOERR | (14<<8))
Private Const P_SQLITE_IOERR_LOCK               As Long = 3850 '(SQLITE_IOERR | (15<<8))
Private Const P_SQLITE_IOERR_CLOSE              As Long = 4106 '(SQLITE_IOERR | (16<<8))
Private Const P_SQLITE_IOERR_DIR_CLOSE          As Long = 4362 '(SQLITE_IOERR | (17<<8))
Private Const P_SQLITE_LOCKED_SHAREDCACHE       As Long = 265  '(SQLITE_LOCKED | (1<<8) )

' Flags For File Open Operations
Private Const P_SQLITE_OPEN_READONLY           As Long = 1       ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_READWRITE          As Long = 2       ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_CREATE             As Long = 4       ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_DELETEONCLOSE      As Long = 8       ' VFS only
Private Const P_SQLITE_OPEN_EXCLUSIVE          As Long = 16      ' VFS only
Private Const P_SQLITE_OPEN_AUTOPROXY          As Long = 32      ' VFS only
Private Const P_SQLITE_OPEN_URI                As Long = 64      ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_MEMORY             As Long = 128     ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_MAIN_DB            As Long = 256     ' VFS only
Private Const P_SQLITE_OPEN_TEMP_DB            As Long = 512     ' VFS only
Private Const P_SQLITE_OPEN_TRANSIENT_DB       As Long = 1024    ' VFS only
Private Const P_SQLITE_OPEN_MAIN_JOURNAL       As Long = 2048    ' VFS only
Private Const P_SQLITE_OPEN_TEMP_JOURNAL       As Long = 4096    ' VFS only
Private Const P_SQLITE_OPEN_SUBJOURNAL         As Long = 8192    ' VFS only
Private Const P_SQLITE_OPEN_MASTER_JOURNAL     As Long = 16384   ' VFS only
Private Const P_SQLITE_OPEN_NOMUTEX            As Long = 32768   ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_FULLMUTEX          As Long = 65536   ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_SHAREDCACHE        As Long = 131072  ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_PRIVATECACHE       As Long = 262144  ' Ok for sqlite3_open_v2()
Private Const P_SQLITE_OPEN_WAL                As Long = 524288  ' VFS only



' Options for Text and Blob binding
Private Const SQLITE_STATIC      As Long = 0
Private Const SQLITE_TRANSIENT   As Long = -1

' System calls
Private Const CP_UTF8 As Long = 65001
#If Win64 Then

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal length As Long)
Private Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr, ByVal cchCount As Long) As LongPtr
Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As LongPtr) As Long
Private Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As LongPtr) As LongPtr
Private Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As LongPtr) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
#Else
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal length As Long)
Private Declare Function lstrcpynW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long, ByVal cchCount As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal pwsString As Long) As Long
Private Declare Function SysAllocString Lib "OleAut32" (ByRef pwsString As Long) As Long
Private Declare Function SysStringLen Lib "OleAut32" (ByVal bstrString As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
#End If
'=====================================================================================
' SQLite StdCall Imports
'-----------------------
#If Win64 Then
' SQLite library version
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
' Database connections
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf16String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long

' Statements
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

' Statement column access (0-based indices)
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString

Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As Double) As Long
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal value As LongLong) As Long
Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

'Backup
Private Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else

' SQLite library version
Private Declare Function sqlite3_libversion Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_libversion@0" () As Long ' PtrUtf8String
' Database connections
Private Declare Function sqlite3_open16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open16@8" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_open_v2@16" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_close@4" (ByVal hDb As Long) As Long
' Database connection error info
Private Declare Function sqlite3_errmsg Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg@4" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errmsg16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errmsg16@4" (ByVal hDb As Long) As Long ' PtrUtf16String
Private Declare Function sqlite3_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_errcode@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_extended_errcode@4" (ByVal hDb As Long) As Long
' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_changes@4" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_total_changes@4" (ByVal hDb As Long) As Long

' Statements
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_prepare16_v2@20" _
    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
Private Declare Function sqlite3_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_step@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_reset Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_reset@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_finalize Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_finalize@4" (ByVal hStmt As Long) As Long

' Statement column access (0-based indices)
Private Declare Function sqlite3_column_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_type@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_name16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_name16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString

Private Declare Function sqlite3_column_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_blob@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_bytes16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_bytes16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_double@8" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_int64@8" (ByVal hStmt As Long, ByVal iCol As Long) As Currency ' UNTESTED ....?
Private Declare Function sqlite3_column_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_text16@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_column_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_column_value@8" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_count@4" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_name@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_parameter_index@8" (ByVal hStmt As Long, ByVal paramName As Long) As Long
Private Declare Function sqlite3_bind_null Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_null@8" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_blob@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_zeroblob@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_bind_double Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_double@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Double) As Long
Private Declare Function sqlite3_bind_int Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_int64@16" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Currency) As Long ' UNTESTED ....?
Private Declare Function sqlite3_bind_text Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_text16 Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_text16@20" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_value Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_bind_value@12" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_clear_bindings@4" (ByVal hStmt As Long) As Long

'Backup
Private Declare Function sqlite3_sleep Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_sleep@4" (ByVal msToSleep As Long) As Long
Private Declare Function sqlite3_backup_init Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_init@16" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_step@8" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_finish@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_remaining@4" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "SQLite3_StdCall" Alias "_sqlite3_stdcall_backup_pagecount@4" (ByVal hBackup As Long) As Long
#End If
'=====================================================================================
' Initialize - load libraries explicitly
#If Win64 Then
    Private hSQLiteLibrary As LongPtr
    Private hSQLiteStdCallLibrary As LongPtr
#Else
    Private hSQLiteLibrary As Long
    Private hSQLiteStdCallLibrary As Long
#End If

Private Property Get Isqlite3_SqliteInitOk() As Long
    Isqlite3_SqliteInitOk = P_SQLITE_INIT_OK
End Property
Private Property Get Isqlite3_SqliteInitError() As Long
    Isqlite3_SqliteInitError = P_SQLITE_INIT_ERROR
End Property
Private Property Get Isqlite3_SqliteInteger() As Long
    Isqlite3_SqliteInteger = P_SQLITE_INTEGER
End Property
Private Property Get Isqlite3_SqliteFloat() As Long
    Isqlite3_SqliteFloat = P_SQLITE_FLOAT
End Property
Private Property Get Isqlite3_SqliteText() As Long
    Isqlite3_SqliteText = P_SQLITE_TEXT
End Property
Private Property Get Isqlite3_SqliteBlob() As Long
    Isqlite3_SqliteBlob = P_SQLITE_BLOB
End Property
Private Property Get Isqlite3_SqliteNull() As Long
    Isqlite3_SqliteNull = P_SQLITE_NULL
End Property
Private Property Get Isqlite3_SqliteOk() As Long
    Isqlite3_SqliteOk = P_SQLITE_OK
End Property
Private Property Get Isqlite3_SqliteError() As Long
    Isqlite3_SqliteError = P_SQLITE_ERROR
End Property
Private Property Get Isqlite3_SqliteInternal() As Long
    Isqlite3_SqliteInternal = P_SQLITE_INTERNAL
End Property
Private Property Get Isqlite3_SqlitePerm() As Long
    Isqlite3_SqlitePerm = P_SQLITE_PERM
End Property
Private Property Get Isqlite3_SqliteAbort() As Long
    Isqlite3_SqliteAbort = P_SQLITE_ABORT
End Property
Private Property Get Isqlite3_SqliteBusy() As Long
    Isqlite3_SqliteBusy = P_SQLITE_BUSY
End Property
Private Property Get Isqlite3_SqliteLocked() As Long
    Isqlite3_SqliteLocked = P_SQLITE_LOCKED
End Property
Private Property Get Isqlite3_SqliteNomem() As Long
    Isqlite3_SqliteNomem = P_SQLITE_NOMEM
End Property
Private Property Get Isqlite3_SqliteReadonly() As Long
    Isqlite3_SqliteReadonly = P_SQLITE_READONLY
End Property
Private Property Get Isqlite3_SqliteInterrupt() As Long
    Isqlite3_SqliteInterrupt = P_SQLITE_INTERRUPT
End Property
Private Property Get Isqlite3_SqliteIoerr() As Long
    Isqlite3_SqliteIoerr = P_SQLITE_IOERR
End Property
Private Property Get Isqlite3_SqliteCorrupt() As Long
    Isqlite3_SqliteCorrupt = P_SQLITE_CORRUPT
End Property
Private Property Get Isqlite3_SqliteNotfound() As Long
    Isqlite3_SqliteNotfound = P_SQLITE_NOTFOUND
End Property
Private Property Get Isqlite3_SqliteFull() As Long
    Isqlite3_SqliteFull = P_SQLITE_FULL
End Property
Private Property Get Isqlite3_SqliteCantopen() As Long
    Isqlite3_SqliteCantopen = P_SQLITE_CANTOPEN
End Property
Private Property Get Isqlite3_SqliteProtocol() As Long
    Isqlite3_SqliteProtocol = P_SQLITE_PROTOCOL
End Property
Private Property Get Isqlite3_SqliteEmpty() As Long
    Isqlite3_SqliteEmpty = P_SQLITE_EMPTY
End Property
Private Property Get Isqlite3_SqliteSchema() As Long
    Isqlite3_SqliteSchema = P_SQLITE_SCHEMA
End Property
Private Property Get Isqlite3_SqliteToobig() As Long
    Isqlite3_SqliteToobig = P_SQLITE_TOOBIG
End Property
Private Property Get Isqlite3_SqliteConstraint() As Long
    Isqlite3_SqliteConstraint = P_SQLITE_CONSTRAINT
End Property
Private Property Get Isqlite3_SqliteMismatch() As Long
    Isqlite3_SqliteMismatch = P_SQLITE_MISMATCH
End Property
Private Property Get Isqlite3_SqliteMisuse() As Long
    Isqlite3_SqliteMisuse = P_SQLITE_MISUSE
End Property
Private Property Get Isqlite3_SqliteNolfs() As Long
    Isqlite3_SqliteNolfs = P_SQLITE_NOLFS
End Property
Private Property Get Isqlite3_SqliteAuth() As Long
    Isqlite3_SqliteAuth = P_SQLITE_AUTH
End Property
Private Property Get Isqlite3_SqliteFormat() As Long
    Isqlite3_SqliteFormat = P_SQLITE_FORMAT
End Property
Private Property Get Isqlite3_SqliteRange() As Long
    Isqlite3_SqliteRange = P_SQLITE_RANGE
End Property
Private Property Get Isqlite3_SqliteNotadb() As Long
    Isqlite3_SqliteNotadb = P_SQLITE_NOTADB
End Property
Private Property Get Isqlite3_SqliteRow() As Long
    Isqlite3_SqliteRow = P_SQLITE_ROW
End Property
Private Property Get Isqlite3_SqliteDone() As Long
    Isqlite3_SqliteDone = P_SQLITE_DONE
End Property
Private Property Get Isqlite3_SqliteIoerrRead() As Long
    Isqlite3_SqliteIoerrRead = P_SQLITE_IOERR_READ
End Property
Private Property Get Isqlite3_SqliteIoerrShortRead() As Long
    Isqlite3_SqliteIoerrShortRead = P_SQLITE_IOERR_SHORT_READ
End Property
Private Property Get Isqlite3_SqliteIoerrWrite() As Long
    Isqlite3_SqliteIoerrWrite = P_SQLITE_IOERR_WRITE
End Property
Private Property Get Isqlite3_SqliteIoerrFsync() As Long
    Isqlite3_SqliteIoerrFsync = P_SQLITE_IOERR_FSYNC
End Property
Private Property Get Isqlite3_SqliteIoerrDirFsync() As Long
    Isqlite3_SqliteIoerrDirFsync = P_SQLITE_IOERR_DIR_FSYNC
End Property
Private Property Get Isqlite3_SqliteIoerrTruncate() As Long
    Isqlite3_SqliteIoerrTruncate = P_SQLITE_IOERR_TRUNCATE
End Property
Private Property Get Isqlite3_SqliteIoerrFstat() As Long
    Isqlite3_SqliteIoerrFstat = P_SQLITE_IOERR_FSTAT
End Property
Private Property Get Isqlite3_SqliteIoerrUnlock() As Long
    Isqlite3_SqliteIoerrUnlock = P_SQLITE_IOERR_UNLOCK
End Property
Private Property Get Isqlite3_SqliteIoerrRdlock() As Long
    Isqlite3_SqliteIoerrRdlock = P_SQLITE_IOERR_RDLOCK
End Property
Private Property Get Isqlite3_SqliteIoerrDelete() As Long
    Isqlite3_SqliteIoerrDelete = P_SQLITE_IOERR_DELETE
End Property
Private Property Get Isqlite3_SqliteIoerrBlocked() As Long
    Isqlite3_SqliteIoerrBlocked = P_SQLITE_IOERR_BLOCKED
End Property
Private Property Get Isqlite3_SqliteIoerrNomem() As Long
    Isqlite3_SqliteIoerrNomem = P_SQLITE_IOERR_NOMEM
End Property
Private Property Get Isqlite3_SqliteIoerrAccess() As Long
    Isqlite3_SqliteIoerrAccess = P_SQLITE_IOERR_ACCESS
End Property
Private Property Get Isqlite3_SqliteIoerrCheckreservedlock() As Long
    Isqlite3_SqliteIoerrCheckreservedlock = P_SQLITE_IOERR_CHECKRESERVEDLOCK
End Property
Private Property Get Isqlite3_SqliteIoerrLock() As Long
    Isqlite3_SqliteIoerrLock = P_SQLITE_IOERR_LOCK
End Property
Private Property Get Isqlite3_SqliteIoerrClose() As Long
    Isqlite3_SqliteIoerrClose = P_SQLITE_IOERR_CLOSE
End Property
Private Property Get Isqlite3_SqliteIoerrDirClose() As Long
    Isqlite3_SqliteIoerrDirClose = P_SQLITE_IOERR_DIR_CLOSE
End Property
Private Property Get Isqlite3_SqliteLockedSharedcache() As Long
    Isqlite3_SqliteLockedSharedcache = P_SQLITE_LOCKED_SHAREDCACHE
End Property
Private Property Get Isqlite3_SqliteOpenReadonly() As Long
    Isqlite3_SqliteOpenReadonly = P_SQLITE_OPEN_READONLY
End Property
Private Property Get Isqlite3_SqliteOpenReadwrite() As Long
    Isqlite3_SqliteOpenReadwrite = P_SQLITE_OPEN_READWRITE
End Property
Private Property Get Isqlite3_SqliteOpenCreate() As Long
    Isqlite3_SqliteOpenCreate = P_SQLITE_OPEN_CREATE
End Property
Private Property Get Isqlite3_SqliteOpenDeleteonclose() As Long
    Isqlite3_SqliteOpenDeleteonclose = P_SQLITE_OPEN_DELETEONCLOSE
End Property
Private Property Get Isqlite3_SqliteOpenExclusive() As Long
    Isqlite3_SqliteOpenExclusive = P_SQLITE_OPEN_EXCLUSIVE
End Property
Private Property Get Isqlite3_SqliteOpenAutoproxy() As Long
    Isqlite3_SqliteOpenAutoproxy = P_SQLITE_OPEN_AUTOPROXY
End Property
Private Property Get Isqlite3_SqliteOpenUri() As Long
    Isqlite3_SqliteOpenUri = P_SQLITE_OPEN_URI
End Property
Private Property Get Isqlite3_SqliteOpenMemory() As Long
    Isqlite3_SqliteOpenMemory = P_SQLITE_OPEN_MEMORY
End Property
Private Property Get Isqlite3_SqliteOpenMainDb() As Long
    Isqlite3_SqliteOpenMainDb = P_SQLITE_OPEN_MAIN_DB
End Property
Private Property Get Isqlite3_SqliteOpenTempDb() As Long
    Isqlite3_SqliteOpenTempDb = P_SQLITE_OPEN_TEMP_DB
End Property
Private Property Get Isqlite3_SqliteOpenTransientDb() As Long
    Isqlite3_SqliteOpenTransientDb = P_SQLITE_OPEN_TRANSIENT_DB
End Property
Private Property Get Isqlite3_SqliteOpenMainJournal() As Long
    Isqlite3_SqliteOpenMainJournal = P_SQLITE_OPEN_MAIN_JOURNAL
End Property
Private Property Get Isqlite3_SqliteOpenTempJournal() As Long
    Isqlite3_SqliteOpenTempJournal = P_SQLITE_OPEN_TEMP_JOURNAL
End Property
Private Property Get Isqlite3_SqliteOpenSubjournal() As Long
    Isqlite3_SqliteOpenSubjournal = P_SQLITE_OPEN_SUBJOURNAL
End Property
Private Property Get Isqlite3_SqliteOpenMasterJournal() As Long
    Isqlite3_SqliteOpenMasterJournal = P_SQLITE_OPEN_MASTER_JOURNAL
End Property
Private Property Get Isqlite3_SqliteOpenNomutex() As Long
    Isqlite3_SqliteOpenNomutex = P_SQLITE_OPEN_NOMUTEX
End Property
Private Property Get Isqlite3_SqliteOpenFullmutex() As Long
    Isqlite3_SqliteOpenFullmutex = P_SQLITE_OPEN_FULLMUTEX
End Property
Private Property Get Isqlite3_SqliteOpenSharedcache() As Long
    Isqlite3_SqliteOpenSharedcache = P_SQLITE_OPEN_SHAREDCACHE
End Property
Private Property Get Isqlite3_SqliteOpenPrivatecache() As Long
    Isqlite3_SqliteOpenPrivatecache = P_SQLITE_OPEN_PRIVATECACHE
End Property
Private Property Get Isqlite3_SqliteOpenWal() As Long
    Isqlite3_SqliteOpenWal = P_SQLITE_OPEN_WAL
End Property

Private Sub Class_Initialize()
    Call InitDirectory
End Sub

Private Sub InitDirectory()
    If P_MASTER_DATABASE_DIRECTORY = "" Then
        Call ISqlite3_CheckOrChangeDir
    Else
        Call ISqlite3_CheckOrChangeDir(P_MASTER_DATABASE_DIRECTORY)
    End If
End Sub

Private Sub ISqlite3_CheckOrChangeDir(Optional MasterDirectory As String = "")
    If MasterDirectory = "" Then
        If CurDir <> ThisWorkbook.Path Then
            ISqlite3_DatabaseDirectory = ThisWorkbook.Path
        End If
    Else
        ISqlite3_DatabaseDirectory = MasterDirectory
    End If
End Sub

Private Property Let ISqlite3_DatabaseDirectory(value As String)
    If P_MASTER_DATABASE_DIRECTORY = "" Then
        pDatabaseDirectory = value
        ChDir pDatabaseDirectory
    Else
        Err.Raise Number:=P_MASTER_DIRECTORY_SET, Description:="Constant variable Master directory set in class module. It must be deleted to overwrite the database directory."
    End If
End Property

Private Property Get ISqlite3_DatabaseDirectory() As String
    If P_MASTER_DATABASE_DIRECTORY = "" Then
        ISqlite3_DatabaseDirectory = pDatabaseDirectory
    Else
        ISqlite3_DatabaseDirectory = "Master directory: " & P_MASTER_DATABASE_DIRECTORY
    End If
End Property

Private Function ISqlite3_SQLite3Initialize(Optional ByVal libDir As String) As Long
    ' A nice option here is to call SetDllDirectory, but that API is only available since Windows XP SP1.
    If libDir = "" Then libDir = ThisWorkbook.Path
    If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
    
    If hSQLiteLibrary = 0 Then
        hSQLiteLibrary = LoadLibrary(libDir + "SQLite3.dll")
        If hSQLiteLibrary = 0 Then
            Debug.Print "SQLite3Initialize Error Loading " + libDir + "SQLite3.dll:", Err.LastDllError
            ISqlite3_SQLite3Initialize = Isqlite3_SqliteInitError
            Exit Function
        End If
    End If
        
    #If Win64 Then
    #Else
    If hSQLiteStdCallLibrary = 0 Then
        hSQLiteStdCallLibrary = LoadLibrary(libDir + "SQLite3_StdCall.dll")
        If hSQLiteStdCallLibrary = 0 Then
            Debug.Print "SQLite3Initialize Error Loading " + libDir + "SQLite3_StdCall.dll:", Err.LastDllError
            ISqlite3_SQLite3Initialize = P_SQLITE_INIT_ERROR
            Exit Function
        End If
    End If
    #End If
    ISqlite3_SQLite3Initialize = Isqlite3_SqliteInitOk
End Function

Private Sub ISqlite3_SQLite3Free()
   Dim refCount As Long
   If hSQLiteStdCallLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteStdCallLibrary)
        hSQLiteStdCallLibrary = 0
        If refCount = 0 Then
            Debug.Print "SQLite3Free Error Freeing SQLite3_StdCall.dll:", refCount, Err.LastDllError
        End If
    End If
    If hSQLiteLibrary <> 0 Then
        refCount = FreeLibrary(hSQLiteLibrary)
        hSQLiteLibrary = 0
        If refCount = 0 Then
            Debug.Print "SQLite3Free Error Freeing SQLite3.dll:", refCount, Err.LastDllError
        End If
    End If
End Sub


'=====================================================================================
' SQLite library version

Private Function ISqlite3_SQLite3LibVersion() As String
    ISqlite3_SQLite3LibVersion = ISqlite3_Utf8PtrToString(sqlite3_libversion())
End Function

'=====================================================================================
' Database connections
Private Function ISqlite3_SQLite3Open(ByVal fileName As String, ByRef dbHandle As LongPtr) As Long
    ISqlite3_SQLite3Open = sqlite3_open16(StrPtr(fileName), dbHandle)
End Function

Private Function ISqlite3_SQLite3OpenV2(ByVal fileName As String, ByRef dbHandle As LongPtr, ByVal flags As Long, ByVal vfsName As String) As Long

    Dim bufFileName() As Byte
    Dim bufVfsName() As Byte
    bufFileName = ISqlite3_StringToUtf8Bytes(fileName)
    If vfsName = Empty Then
        ISqlite3_SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, 0)
    Else
        bufVfsName = ISqlite3_StringToUtf8Bytes(vfsName)
        ISqlite3_SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, flags, VarPtr(bufVfsName(0)))
    End If

End Function

Private Function ISqlite3_SQLite3Close(ByVal dbHandle As LongPtr) As Long
    ISqlite3_SQLite3Close = sqlite3_close(dbHandle)
End Function

'=====================================================================================
' Error information

Private Function ISqlite3_SQLite3ErrMsg(ByVal dbHandle As LongPtr) As String
    ISqlite3_SQLite3ErrMsg = ISqlite3_Utf8PtrToString(sqlite3_errmsg(dbHandle))
End Function

Private Function ISqlite3_SQLite3ErrCode(ByVal dbHandle As LongPtr) As Long
    ISqlite3_SQLite3ErrCode = sqlite3_errcode(dbHandle)
End Function

Private Function ISqlite3_SQLite3ExtendedErrCode(ByVal dbHandle As LongPtr) As Long
    ISqlite3_SQLite3ExtendedErrCode = sqlite3_extended_errcode(dbHandle)
End Function

'=====================================================================================
' Change Counts

Private Function ISqlite3_SQLite3Changes(ByVal dbHandle As LongPtr) As Long
    ISqlite3_SQLite3Changes = sqlite3_changes(dbHandle)
End Function

Private Function ISqlite3_SQLite3TotalChanges(ByVal dbHandle As LongPtr) As Long
    ISqlite3_SQLite3TotalChanges = sqlite3_total_changes(dbHandle)
End Function

'=====================================================================================
' Statements

Private Function ISqlite3_SQLite3PrepareV2(ByVal dbHandle As LongPtr, ByVal sql As String, ByRef stmtHandle As LongPtr) As Long
    ' Only the first statement (up to ';') is prepared. Currently we don't retrieve the 'tail' pointer.
    ISqlite3_SQLite3PrepareV2 = sqlite3_prepare16_v2(dbHandle, StrPtr(sql), Len(sql) * 2, stmtHandle, 0)
End Function

Private Function ISqlite3_SQLite3Step(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3Step = sqlite3_step(stmtHandle)
End Function

Private Function ISqlite3_SQLite3Reset(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3Reset = sqlite3_reset(stmtHandle)
End Function

Private Function ISqlite3_SQLite3Finalize(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3Finalize = sqlite3_finalize(stmtHandle)
End Function

'=====================================================================================
' Statement column access (0-based indices)

Private Function ISqlite3_SQLite3ColumnCount(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3ColumnCount = sqlite3_column_count(stmtHandle)
End Function

Private Function ISqlite3_SQLite3ColumnType(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
    ISqlite3_SQLite3ColumnType = sqlite3_column_type(stmtHandle, ZeroBasedColIndex)
End Function

Private Function ISqlite3_SQLite3ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
    ISqlite3_SQLite3ColumnName = ISqlite3_Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
End Function

Private Function ISqlite3_SQLite3ColumnDouble(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Double
    ISqlite3_SQLite3ColumnDouble = sqlite3_column_double(stmtHandle, ZeroBasedColIndex)
End Function

Private Function ISqlite3_SQLite3ColumnInt32(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
    ISqlite3_SQLite3ColumnInt32 = sqlite3_column_int(stmtHandle, ZeroBasedColIndex)
End Function

Private Function ISqlite3_SQLite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
    ISqlite3_SQLite3ColumnText = ISqlite3_Utf8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
End Function

Private Function ISqlite3_SQLite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date
    ISqlite3_SQLite3ColumnDate = ISqlite3_FromJulianDay(sqlite3_column_double(stmtHandle, ZeroBasedColIndex))
End Function

Private Function ISqlite3_SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As LongPtr

    Dim length As Long
    Dim buf() As Byte
    
    ptr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
    length = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim buf(length - 1)
    RtlMoveMemory VarPtr(buf(0)), ptr, length
    ISqlite3_SQLite3ColumnBlob = buf
End Function
'=====================================================================================
' Statement bindings

Private Function ISqlite3_SQLite3BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As String) As Long
    ISqlite3_SQLite3BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(value), -1, SQLITE_TRANSIENT)
End Function

Private Function ISqlite3_SQLite3BindDouble(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Double) As Long
    ISqlite3_SQLite3BindDouble = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, value)
End Function

Private Function ISqlite3_SQLite3BindInt32(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Long) As Long
    ISqlite3_SQLite3BindInt32 = sqlite3_bind_int(stmtHandle, OneBasedParamIndex, value)
End Function

Private Function ISqlite3_SQLite3BindDate(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal value As Date) As Long
    ISqlite3_SQLite3BindDate = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, ISqlite3_ToJulianDay(value))
End Function

Private Function ISqlite3_SQLite3BindBlob(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByRef value() As Byte) As Long
    Dim length As Long
    length = UBound(value) - LBound(value) + 1
    ISqlite3_SQLite3BindBlob = sqlite3_bind_blob(stmtHandle, OneBasedParamIndex, VarPtr(value(0)), length, SQLITE_TRANSIENT)
End Function

Private Function ISqlite3_SQLite3BindNull(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As Long
    ISqlite3_SQLite3BindNull = sqlite3_bind_null(stmtHandle, OneBasedParamIndex)
End Function

Private Function ISqlite3_SQLite3BindParameterCount(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3BindParameterCount = sqlite3_bind_parameter_count(stmtHandle)
End Function

Private Function ISqlite3_SQLite3BindParameterName(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As String
    ISqlite3_SQLite3BindParameterName = ISqlite3_Utf8PtrToString(sqlite3_bind_parameter_name(stmtHandle, OneBasedParamIndex))
End Function

Private Function ISqlite3_SQLite3BindParameterIndex(ByVal stmtHandle As LongPtr, ByVal paramName As String) As Long
    Dim buf() As Byte
    buf = ISqlite3_StringToUtf8Bytes(paramName)
    ISqlite3_SQLite3BindParameterIndex = sqlite3_bind_parameter_index(stmtHandle, VarPtr(buf(0)))
End Function

Private Function ISqlite3_SQLite3ClearBindings(ByVal stmtHandle As LongPtr) As Long
    ISqlite3_SQLite3ClearBindings = sqlite3_clear_bindings(stmtHandle)
End Function

'=====================================================================================
' Backup
Private Function ISqlite3_SQLite3Sleep(ByVal timeToSleepInMs As Long) As Long
    ISqlite3_SQLite3Sleep = sqlite3_sleep(timeToSleepInMs)
End Function

Private Function ISqlite3_SQLite3BackupInit(ByVal dbHandleDestination As LongPtr, ByVal destinationName As String, ByVal dbHandleSource As LongPtr, ByVal sourceName As String) As LongPtr
    Dim bufDestinationName() As Byte
    Dim bufSourceName() As Byte
    bufDestinationName = ISqlite3_StringToUtf8Bytes(destinationName)
    bufSourceName = ISqlite3_StringToUtf8Bytes(sourceName)
    ISqlite3_SQLite3BackupInit = sqlite3_backup_init(dbHandleDestination, VarPtr(bufDestinationName(0)), dbHandleSource, VarPtr(bufSourceName(0)))
End Function

Private Function ISqlite3_SQLite3BackupFinish(ByVal backupHandle As LongPtr) As Long
    ISqlite3_SQLite3BackupFinish = sqlite3_backup_finish(backupHandle)
End Function

Private Function ISqlite3_SQLite3BackupStep(ByVal backupHandle As LongPtr, ByVal numberOfPages) As Long
    ISqlite3_SQLite3BackupStep = sqlite3_backup_step(backupHandle, numberOfPages)
End Function

Private Function ISqlite3_SQLite3BackupPageCount(ByVal backupHandle As LongPtr) As Long
    ISqlite3_SQLite3BackupPageCount = sqlite3_backup_pagecount(backupHandle)
End Function

Public Function ISqlite3_SQLite3BackupRemaining(ByVal backupHandle As LongPtr) As Long
    ISqlite3_SQLite3BackupRemaining = sqlite3_backup_remaining(backupHandle)
End Function

' String Helpers
Private Function ISqlite3_Utf8PtrToString(ByVal pUtf8String As LongPtr) As String
    Dim buf As String
    Dim cSize As Long
    Dim RetVal As Long
    
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
    ' cSize includes the terminating null character
    If cSize <= 1 Then
        ISqlite3_Utf8PtrToString = ""
        Exit Function
    End If
    
    ISqlite3_Utf8PtrToString = String(cSize - 1, "*") ' and a termintating null char.
    RetVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(ISqlite3_Utf8PtrToString), cSize)
    If RetVal = 0 Then
        Debug.Print "Utf8PtrToString Error:", Err.LastDllError
        Exit Function
    End If
End Function

Private Function ISqlite3_StringToUtf8Bytes(ByVal str As String) As Variant
    Dim bSize As Long
    Dim RetVal As Long
    Dim buf() As Byte
    
    bSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, 0, 0, 0, 0)
    If bSize = 0 Then
        Exit Function
    End If
    
    ReDim buf(bSize)
    RetVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), -1, VarPtr(buf(0)), bSize, 0, 0)
    If RetVal = 0 Then
        Debug.Print "StringToUtf8Bytes Error:", Err.LastDllError
        Exit Function
    End If
    ISqlite3_StringToUtf8Bytes = buf
End Function

Private Function ISqlite3_Utf16PtrToString(ByVal pUtf16String As LongPtr) As String
    Dim StrLen As Long
    StrLen = lstrlenW(pUtf16String)
    ISqlite3_Utf16PtrToString = String(StrLen, "*")
    lstrcpynW StrPtr(ISqlite3_Utf16PtrToString), pUtf16String, StrLen
End Function

' Date Helpers
Private Function ISqlite3_ToJulianDay(oleDate As Date) As Double
    ISqlite3_ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Private Function ISqlite3_FromJulianDay(julianDay As Double) As Date
    ISqlite3_FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function


