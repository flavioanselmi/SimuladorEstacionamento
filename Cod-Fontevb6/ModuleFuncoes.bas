Attribute VB_Name = "ModFuncoes"


Public Function funcIntGerarNumRND(vMax, vMin As Integer) As Integer

    'retorna um n�mero inteiro rand�mico dentro de um intervalo definido
    Randomize
    funcIntGerarNumRND = Int((Rnd() * (vMax - vMin + 1)) + vMin)

End Function
