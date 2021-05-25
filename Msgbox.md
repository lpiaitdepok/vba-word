```VisualBasic
MsgBox ( Prompt [,icons+buttons ] [,title ] )
VariableInteger = MsgBox ( prompt [, icons+ buttons] [,title] )
```

# Icons:
| Constant | Nilai | Keterangan | 
| :--- | :--- | :--- | 
| vbCritical | 16 | Icon Critical | 
| vbQuestion | 32 |  Icon Tanda Tanya | 
| vbExclamation | 48 | Icon peringatan (warning) | 
| vbInformation | 64 | Icon Informasi | 

# Tombol
| Constant           | Nilai | Keterangan                                 |
| :----------------- | :---- | :----------------------------------------- |
| vbOkOnly           | 0     | Menampilkan tombol OK                      |
| vbOkCancel         | 1     | Menampilkan tombol  OK dan Cancel          |
| vbAbortRetryIgnore | 2     | Menampilkan tombol Abort, Retry dan Ignore |
| vbYesNoCancel      | 3     | Menampilkan tombol Yes, No dan Cancel      |
| vbYesNo            | 4     | Menampilkan tombol Yes dan No              |
| vbRetryCancel      | 5     | Menampilkan tombol Retry dan Cancel        |

# Nilai yang dikembalikan (hasil)
| Constant | Nilai | Keterangan    |
| :------- | :---- | :------------ |
| vbOk     | 1     | Tombol Ok     |
| vbCancel | 2     | Tombol Cancel |
| vbAbort  | 3     | Tombol Abort  |
| vbRetry  | 4     | Tombol Retry  |
| vbIgnore | 5     | Tombol Ignore |
| vbYes    | 6     | Tombol Yes    |
| vbNo     | 7     | Tombol No     |
|          |       |               |
