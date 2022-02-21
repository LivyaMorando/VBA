'Exemplo 2 - considerando as faltas e as aulas
'Frequencia maior ou igual a 75% - aprovado por frequencia
'Frequencia menor que 75% - reprovado por frequencia

'se a nota for maior ou igual a 6 ele está aprovado
'se a nota for igual a 4 e menor que 6 ele está de exame
'se a nota for menor que 4 ele foi reprovado

'dividindo as faltas pelas aulas tenho a % de faltas, fazendo 1 - esse valor tenho a frquencia

Sub ProcessaNotaFreq ()
    nota = Range ("E3").value

    Dim freq as duble 
    freq = 1 -( Range ("D3").value / range("C3").value )

    if freq < 0.75 then 
    range("F3").value="Reprovado por frequencia"
    Range("F3").interior.colorindex = 3
    else 
 if nota >= 6 and nota=< 10 then 
    Range("F3").value="Aprovado"
    Range("F3").interior.colorindex = 4
  Elseif nota < 4 then     'se não se
    Range("F3").value="Reprovado"
    Range("F3").interior.colorindex = 3
  Elseif nota >= 4 and nota < 6 then
    Range("F3").value="Exame"
    Range("F3").interior.colorindex = 6
  Else 
    Range("F3").value = "Erro"
    Range("F3").interior.colorindex = 2
  End if


End Sub


Sub ProcessaNota6 ()
  Dim nota as duble
  nota = Range("C3").Value

  if nota >= 6 and nota=< 10 then 
    Range("D3").value="Aprovado"
    Range("D3").interior.colorindex = 4
  Elseif nota < 4 then     'se não se
    Range("D3").value="Reprovado"
    Range("D3").interior.colorindex = 3
  Elseif nota >= 4 and nota < 6 then
    Range("D3").value="Exame"
    Range("D3").interior.colorindex = 6
  Else 
    Range("D3").value = "Erro"
    Range("D3").interior.colorindex = 2
  End if

  End sub 

