'=========================================================================
' Arrays - Join Function
'=========================================================================

' Error Handling
option explicit

' Variables
Dim arrWebsite(2)

arrWebsite(0) = "www"
arrWebsite(1) = "CodingGears"
arrWebsite(2) = "io"

Dim siteUrl
siteUrl = Join(arrWebsite, ".")

MsgBox siteUrl