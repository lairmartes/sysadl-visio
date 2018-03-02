VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagramCommunicationSysAdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pFirstStructure As DiagramStructureSysAdl
    Private pSecondStructure As DiagramStructureSysAdl
    Private pCommunicationShapeController As shapeController
    
    Public Sub Init(ByVal aFirstStructure As DiagramStructureSysAdl, _
                    ByVal aSecondStructure As DiagramStructureSysAdl, _
                    ByVal aCommunicationShapeController As shapeController)
        
        Set pFirstStructure = aFirstStructure
        Set pSecondStructure = aSecondStructure
        Set pCommunicationShapeController = aCommunicationShapeController
    
    End Sub
    
    Public Property Get Id() As String
    
        Dim Result As String
        
        Result = pCommunicationShapeController.GetShapeUniqueId
        
        Id = Result
    
    End Property
    
    Public Property Get FirstStructure() As DiagramStructureSysAdl
    
        Set FirstStructure = pFirstStructure
    
    End Property

    Public Property Get SecondStrucutre() As DiagramStructureSysAdl
    
        Set SecondStrucutre = pSecondStructure

    End Property

    Public Property Get CommunicationShapeController() As shapeController
    
        Set CommunicationShapeController = pCommunicationShapeController
        
    End Property
    
    Public Property Get CommunicationType() As String
    
        Dim Result As String
        
        Result = CalculateCommunicationType()
    
        CommunicationType = Result
        
    End Property
    
    Private Function CalculateCommunicationType() As String
    
        Dim Result As String
        
        Dim firstStructureType As String
        Dim secondStructureType As String
        
        firstStructureType = pFirstStructure.StructureType
        secondStructureType = pSecondStructure.StructureType
        

        If ((firstStructureType = sysAdlStructureInterface And secondStructureType = sysAdlStructureProtocol) _
        Or (firstStructureType = sysAdlStructureProtocol And secondStructureType = sysAdlStructureInterface)) Then
        
            Result = sysAdlCommunicationConversation
            
        ElseIf (firstStructureType = sysAdlStructureDeposit And secondStructureType = sysAdlStructureSingleElement) Then

             If (pSecondStructure.GetSingleElementType = sysAdlTypeSetSystem) Then
             
                Result = sysAdlCommunicationConversation
                
             End If
             
        ElseIf (firstStructureType = sysAdlStructurePort And secondStructureType = sysAdlStructureSingleElement) Then
        
            If (pSecondStructure.GetSingleElementType = sysAdlTypeSetNode) Then
            
                Result = sysAdlCommunicationNet
                
            End If
            
        ElseIf (firstStructureType = sysAdlStructureDevice And secondStructureType = sysAdlStructureSingleElement) Then
        
            If (pSecondStructure.GetSingleElementType = sysAdlTypeSetNode) Then
                
                Result = sysAdlCommunicationInstallation
                
            End If
            
            
        ElseIf (firstStructureType = sysAdlStructureSingleElement) Then
        
            If (pFirstStructure.GetSingleElementType = sysAdlTypeSetSystem And secondStructureType = sysAdlStructureDeposit) Then
            
                Result = sysAdlCommunicationConversation
                
            ElseIf (pFirstStructure.GetSingleElementType = sysAdlTypeSetNode And secondStructureType = sysAdlStructurePort) Then
                
                Result = sysAdlCommunicationConversation
                
            ElseIf (pFirstStructure.GetSingleElementController = sysAdlTypeSetNode And secondStructureType = sysAdlStructureDevice) Then
            
                Result = sysAdlCommunicationInstallation
                
            End If
             
        End If
        
        CalculateCommunicationType = Result
        
        
    End Function
    
    Public Function GetStructureByType(ByVal StructureType As String) As DiagramStructureSysAdl
    
        Dim Result As DiagramStructureSysAdl
    
        If pFirstStructure.StructureType = StructureType Then
        
            Set Result = pFirstStructure
            
        ElseIf pSecondStructure.StructureType = StructureType Then
        
            Set Result = pSecondStructure
            
        End If
        
        Set GetStructureByType = Result
        
    End Function
    
    
    Public Function GetConversationType() As String
    
        Dim Result As String
        
        Result = sysAdlStringConstantsEmpty
        
        If pFirstStructure.StructureType = sysAdlStructureProtocol Or pSecondStructure.StructureType = sysAdlStructureProtocol Then
        
            Result = sysAdlStructureProtocol
            
        ElseIf pFirstStructure.StructureType = sysAdlStructureDeposit Or pSecondStructure.StructureType = sysAdlStructureDeposit Then
        
            Result = sysAdlStructureDeposit
            
        End If
        
        GetConversationType = Result
    
    End Function
