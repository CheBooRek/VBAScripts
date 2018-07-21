Attribute VB_Name = "GetRate"
Function GetRate(ByVal CurrencyName As String, ByVal RateDate As Date) As Double
    ' функция возвращает курс валюты CurrencyName на дату RateDate
    ' в случае ошибки возвращает (неверная дата или название валюты) возвращается 0
    On Error Resume Next
    CurrencyName = UCase(CurrencyName): If Len(CurrencyName) <> 3 Then Exit Function
    Set xmldoc = CreateObject("Msxml.DOMDocument"): xmldoc.async = False
    url_request = "http://www.cbr.ru/scripts/XML_daily.asp?date_req=" + Format(RateDate, "dd\/mm\/yyyy")
    
    If xmldoc.Load(url_request) <> True Then Exit Function ' Запрос к серверу ЦБР
    
    'Обработка полученного ответа
    Set nodeList = xmldoc.SelectNodes("ValCurs"): Set xmlNode = nodeList.Item(0).CloneNode(True)
    Set node_attr = xmlNode.Attributes(0): strDate = node_attr.Value
    Set nodeList = xmldoc.SelectNodes("*/Valute")
    For i = 0 To nodeList.Length - 1 'Поиск нужной валюты
        Set xmlNode = nodeList.Item(i).CloneNode(True)
        If xmlNode.ChildNodes(1).Text = CurrencyName Then
            CurrencyRate = CDbl(xmlNode.ChildNodes(4).Text)
            divisor = Val(xmlNode.ChildNodes(2).Text)
            GetRate = CurrencyRate / divisor
            Exit Function
        End If
    Next i
End Function

