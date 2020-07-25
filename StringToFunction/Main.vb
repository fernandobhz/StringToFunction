Imports System.CodeDom.Compiler
Imports System.Text
Imports System.Reflection

Public Module Main

    Function Eval(vbCode As String, ParamArray params() As Object)
        Dim cp As CompilerParameters = New CompilerParameters
        cp.ReferencedAssemblies.Add("system.dll")
        cp.ReferencedAssemblies.Add("system.xml.dll")
        cp.ReferencedAssemblies.Add("system.data.dll")
        cp.CompilerOptions = "/t:library"
        cp.GenerateInMemory = True


        Dim sb As StringBuilder = New StringBuilder("")
        sb.Append("Imports System" & vbCrLf)
        sb.Append("Imports System.Xml" & vbCrLf)
        sb.Append("Imports System.Data" & vbCrLf)
        sb.Append("Imports System.Data.SqlClient" & vbCrLf)
        sb.Append("Namespace PAB  " & vbCrLf)
        sb.Append("Class PABLib " & vbCrLf)

        sb.Append("public function  EvalCode(x) as Object " & vbCrLf)
        sb.Append(vbCode & vbCrLf)
        sb.Append("End Function " & vbCrLf)
        sb.Append("End Class " & vbCrLf)
        sb.Append("End Namespace" & vbCrLf)

        Dim c As VBCodeProvider = New VBCodeProvider
        Dim cr = c.CompileAssemblyFromSource(cp, sb.ToString())
        Dim a = cr.CompiledAssembly
        Dim o As Object
        Dim mi As MethodInfo
        o = a.CreateInstance("PAB.PABLib")
        Dim t As Type = o.GetType()
        mi = t.GetMethod("EvalCode")
        Dim s As Object
        s = mi.Invoke(o, params)
        Return s
    End Function

    Sub Main()
        Dim X = New With {.a = "abc", .b = "abc"}

        Dim sb As New StringBuilder
        sb.AppendLine("if x.a = x.b then")
        sb.AppendLine("  return true")
        sb.AppendLine("else")
        sb.AppendLine("  return false")
        sb.AppendLine("end if")

        Dim Result As Boolean = Eval(sb.ToString, X)

        MsgBox(Result)
    End Sub

End Module
