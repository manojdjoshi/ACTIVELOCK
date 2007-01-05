Public Interface iLicense
    '
    ' Define Public Methods
    '
    Function Validate(ByVal RequestData As LockInfo) As String
    Function Acquire(ByVal RequestData As LockInfo) As Boolean
    Function Verify(ByVal RequestData As LockInfo) As Boolean
    Sub Release(ByVal RequestData As LockInfo)
    '
    ' Done
    '
End Interface
