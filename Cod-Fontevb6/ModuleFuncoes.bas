Attribute VB_Name = "ModFuncoes"


Public Function funcIntGerarNumRND(vMax, vMin As Integer) As Integer

    'retorna um número inteiro randômico dentro de um intervalo definido
    Randomize
    funcIntGerarNumRND = Int((Rnd() * (vMax - vMin + 1)) + vMin)

End Function
