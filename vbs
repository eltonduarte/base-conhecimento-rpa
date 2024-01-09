' Retorna o primeiro dia do mês anterior
Function DataInicial()
  DataInicial = DateSerial(Year(Date), Month(Date) - 1, 1)
  DataInicial = Replace(DataInicial, "/", ".")
End function

' Retorna o último dia do mês anterior
Function DataFinal()
  DataFinal = DateSerial(Year(Date), Month(Date), 0)
  DataFinal = Replace(DataFinal, "/", ".")
End function
