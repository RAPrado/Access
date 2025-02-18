# Limpar caracteres usando regex / Clear character using regex  
Parâmetro :<br>
**Conteudo -** recebe texto a ser limpo / receive a text to be clear.

**Retorno -** conteudo apenas com os caracteres aceitos / return a text just with de characters allowed

```vba
Function Limpar(Conteudo As String) As String
    Dim Regex As Object
    Dim matches As Object
    Dim match As Object
    Dim Retorno As String

    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = "[A-Za-z 0-9ÁáÉéÍíÓóÚúÀàÈèÌìÒòÙùÂâÊêÎîÔôÛûÃãÕõÄäËëÏïÖöÜüçÇýÝ!#$€%&/{}\[\]()""=?«»*+ºª<>,;.:\-_\\|]|\r\n" 'Padrões aceitos. E para o padrão \r\n significa Enter

    Regex.Global = True ' Encontra todas as ocorrências

    Set matches = Regex.Execute(Conteudo)

    For Each match In matches
        Retorno = Retorno & match.value
    Next match

    Limpar = Retorno

    Set Regex = Nothing
    Set matches = Nothing
End Function
```
