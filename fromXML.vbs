'  入力ファイル
Dim inputFileName
inputFileName = "sample_input.xml"

' XMLの読み込み準備を行う
Dim xr
Set xr = New XmlReader
xr.LoadXmlFile (inputFileName)

' XMLよりデータを読み込む
Dim memberList()
Call xr.GetMemberList(memberList)

' 出力 文字列
Dim outputStr

' 取得結果をセルに出力する
' If Sgn(memberList) <> 0 Then

    Dim i
    For i = 0 To UBound(memberList)
        outputStr = outputStr & "id=" & memberList(i).id & ", "
        outputStr = outputStr & "name=" & memberList(i).membername & ", "
        outputStr = outputStr & "age=" & memberList(i).age & vbCrLf
    Next
' End If
MsgBox outputStr, vbInformation, "result"

Set xr = Nothing

Class Member
  Sub Class_Initialize()
  End Sub
  Public id
  Public membername
  Public age
End Class

Class XmlReader
  ' DOM
  Private xmlDocument

  ' コンストラクタ
  Sub Class_Initialize()
  End Sub

  ' XMLをDOMオブジェクトにロードする
  Sub LoadXmlFile(ByVal fileName)
      ' MSXMLオブジェクトを生成
      Set xmlDocument = Nothing
      Set xmlDocument = WScript.CreateObject("MSXML2.DOMDocument")
      xmlDocument.load(fileName)
  End Sub

  ' メンバリストを取得する
  Function GetMemberList(ByRef memberList())
      Dim membersNode
      Dim memberNode
      Dim memberAttribute

      ' XMLのmemberノードを取得する
      Set membersNode = xmlDocument.SelectSingleNode("//members")
      Dim i
      i = 0
      For Each memberNode In membersNode.childNodes
        ReDim Preserve memberList(i)
          Dim newmember
          Set newmember = New Member
          ' idの属性値を取得する
          For Each memberAttribute In memberNode.Attributes
              If memberAttribute.name = "id" Then
                  newmember.id = memberAttribute.Value
              End If
          Next

          ' memberの子要素を取得する
          Dim childNode
          For Each childNode In memberNode.childNodes
              ' name要素の値を取得する
              If childNode.nodeName = "name" Then
                  newmember.membername = childNode.text
              End If
              ' age要素の値を取得する
              If childNode.nodeName = "age" Then
                  newmember.age = childNode.text
              End If
          Next
          Set memberList(i) = newmember


          i = i + 1
      Next

  End Function

  ' デストラクタ
  Public Sub Class_Terminate()
      If Not xmlDocument Is Nothing Then Set xmlDocument = Nothing
  End Sub
End Class
