Condicionais - If, then, elseif e else

'Aluno
'Nota que ele tirou
'vamos escrever uma sub que vai definir a aprovação

'se a nota for maior ou igual a 6 ele está aprovado
'se a nota for igual a 4 e menor que 6 ele está de exame
'se a nota for menor que 4 ele foi reprovado

Sub ProcessaNota ()
  Dim nota as duble 'declaro uma variavel nota usando o Dim e declaro que ela é do tipo duble (que tem valores com virgula)
  nota = Range("C3").Value

  if nota >= 6 then 
    Range("D3").value="Aprovado"
  end if

   if nota < 4 then 
    Range("D3").value="Reprovado"
  end if

 if nota <= 4 and nota < 6 then 
    Range("D3").value="Exame"
  end if

End Sub

'Verificando uma forma mais avançada de fazer o caso acima

Sub ProcessaNota2 ()
  Dim nota as duble 
  nota = Range("C3").Value

  if nota >= 6 then 
    Range("D3").value="Aprovado"
  else 
     if nota < 4 then 
    Range("D3").value="Reprovado"
  Else
    Range("D3").value="Exame"
  End If

  End sub

'Verificando uma forma mais avançada de fazer o caso acima

Sub ProcessaNota3 ()
  Dim nota as duble
  nota = Range("C3").Value

  if nota >= 6 then 
    Range("D3").value="Aprovado"
  Elseif nota < 4 then     'se não se
    Range("D3").value="Reprovado"
  Else
    Range("D3").value="Exame"
  End if

  End sub

'Testando se a nota é maior do que 10

Sub ProcessaNota4 ()
  Dim nota as duble
  nota = Range("C3").Value

  if nota >= 6 and nota=< 10 then 
    Range("D3").value="Aprovado"
  Elseif nota < 4 then     'se não se
    Range("D3").value="Reprovado"
  Elseif nota >= 4 and nota < 6 then
    Range("D3").value="Exame"
  Else 
    Range("D3").value = "Erro"
  End if

  End sub


  'Colorindo as celulas de acordo com a devolutiva
Sub ProcessaNota5 ()
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