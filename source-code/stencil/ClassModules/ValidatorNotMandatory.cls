VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorNotMandatory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IValidator


Public Function IValidator_Validate(ByVal value As String) As String

    Dim Result As String
    
    Result = sysAdlStringConstantsEmpty ' it always will be valid
    
    IValidator_Validate = Result

End Function
