Olá mundo
Inserir um múdulo na pasta 1

Sub OlaMundo () #sequencia de passos que eu quero que ele execute ao chamar o OlaMundo
    Range("A1").Value = "Olá Mundo" #Range retorna uma célula e quando coloco o .value é o valor daquela célula - estou acessando esse valor
End Sub

#Podemos inserir um botão para inserir nossa Sub

#vamos fazer outra sub para apagar

Sub ApagaTexto ()
    Range("A1").value = ""
End Sub

#Suponhamos que eu queira informar um nome e depois usar esse nome em outro lugar

Sub Nome ()
    Range("B2").value = "Olá " & Range("B1").value
End Sub

Sub ApagaNome()
    Range("B2").value = ""
End Sub


