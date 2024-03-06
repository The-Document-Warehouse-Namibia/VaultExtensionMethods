Option Explicit

Const DELIMITER = ";"

Const MF_OBJECT_TYPE = "GetObjectTypeIDByAlias"
Const MF_CLASS = "GetClassIDByAlias"
Const MF_WORKFLOW = "GetWorkflowIDByAlias"
Const MF_WORKFLOW_STATE = "GetWorkflowStateIDByAlias"
Const MF_PROPERTYDEF = "GetPropertyDefIDByAlias"

Public SplitStrings
SplitStrings = Split(Input, DELIMITER)

Public FunctionName : FunctionName = SplitStrings(0)
Public PropertyAlias : PropertyAlias = SplitStrings(1)

Select Case FunctionName
	Case MF_OBJECT_TYPE
		Output = GetObjectTypeIDByAlias(PropertyAlias)
	Case MF_CLASS
		Output = GetClassIDByAlias(PropertyAlias)
	Case MF_PROPERTYDEF
		Output = GetPropertyDefIDByAlias(PropertyAlias)
End Select

' Get object type id by alias helper.
Private Function GetObjectTypeIDByAlias(objectTypeAlias)
  Dim objectTypeID
  objectTypeID = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias(objectTypeAlias)

  If (objectTypeID = -1) Then
    ThrowCustomException NotFoundException("Object Type", objectTypeAlias)
  End If

  GetObjectTypeIDByAlias = objectTypeID
End Function

' Get class id by alias helper.
Private Function GetClassIDByAlias(classAlias)
  Dim classID
  classID = Vault.ClassOperations.GetObjectClassIDByAlias(classAlias)

  If(classID = -1) Then
    ThrowCustomException NotFoundException("Class", classAlias)
  End If

  GetClassIDByAlias = classID
End Function

' Get workflow id by alias helper.
Private Function GetWorkflowIDByAlias(workflowAlias)
  Dim workflowId
  workflowId = Vault.WorkflowOperations.GetWorkflowIDByAlias(workflowAlias)

  If (workflowId = -1) Then
    ThrowCustomException NotFoundException("Workflow", workflowAlias)
  End If

  GetWorkflowIDByAlias = workflowId
End Function

' Get workflow state id by alias helper.
Private Function GetWorkflowStateIDByAlias(workflowStateAlias)
    Dim workflowStateId
    workflowStateId = Vault.WorkflowOperations.GetWorkflowStateIDByAlias(workflowStateAlias)

    If (workflowStateId = -1) Then
      ThrowCustomException NotFoundException("Workflow Satate", workflowStateAlias)
    End If

    GetWorkflowStateIDByAlias = workflowStateId
End Function

' Get property def id by alias helper.
Private Function GetPropertyDefIDByAlias(propertyAlias)
  Dim propertyDefinitionID
  propertyDefinitionID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(propertyAlias)

  If (propertyDefinitionID = -1) Then
    ThrowCustomException NotFoundException("Property", propertyAlias)
  End If

  GetPropertyDefIDByAlias = propertyDefinitionID
End Function

Private Sub ThrowCustomException(exceptionMessage)
  Err.Raise MFScriptCancel, exceptionMessage
End Sub

Private Function NotFoundException(objType, objAlias)
  NotFoundException = objType & " with alias '" & objAlias & "' does not exist or there is more than one " & LCase(objType) & " with the same alias." & vbCrLf &_
            "Contact your M-Files administrator."
End Function