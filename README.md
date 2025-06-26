# 📊 Dashboard de Vendas — Xbox Game Pass

Este projeto foi desenvolvido como parte de um desafio de Excel e análise de dados dp bootcamp da DIO - Santander - Excel com Inteligência Artificial. O objetivo foi transformar dados brutos de vendas em um dashboard visual, interativo e funcional, permitindo análise clara e tomada de decisão baseada em dados.

## 🧩 Funcionalidades

- **Resumo de Vendas:** Total de vendas, valor total, e valores específicos por add-ons como EA Play e Minecraft.
- **Filtros Interativos:** Segmentações por plano, auto renovação, e passes adicionais.
- **Gráficos Dinâmicos:** Barras, colunas e gráfico de pizza para facilitar a análise por plano, mês e tipo de assinatura.
- **Botão de Tela Cheia:** Ativa e desativa o modo imersivo no Excel com um clique.
- **Botão "Limpar Filtros":** Restaura os filtros para o estado inicial automaticamente com macro VBA.

## 💻 Tecnologias Utilizadas

- **Excel 365**
- **Gráficos Dinâmicos**
- **Segmentações de Dados (Slicers)**
- **VBA para automações**

## 🔧 Macros VBA Utilizadas

### 🧼 Limpar Filtros

```vba
Sub LIMPAR_FILTROS()
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Plan").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Auto_Renewal").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_EA_Play_Season_Pass").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Minecraft_Season_Pass").ClearManualFilter
End Sub
```
### 🖥️ Alternar Tela Cheia

```vba
Sub AlternarTelaCheia()
    With Application
        If telaCheiaAtivada = False Then
            .DisplayFullScreen = True
            .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
            .DisplayFormulaBar = False
            .DisplayStatusBar = False
            telaCheiaAtivada = True
            With ActiveWindow
                .DisplayWorkbookTabs = False
                .DisplayHeadings = False
            End With
        Else
            .DisplayFullScreen = False
            .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
            .DisplayFormulaBar = False
            .DisplayStatusBar = False
            telaCheiaAtivada = False
            With ActiveWindow
                .DisplayWorkbookTabs = True
                .DisplayHeadings = False
            End With
        End If
    End With
End Sub
```

### 📁 Como Usar
- Baixe ou clone este repositório.
- Abra o arquivo Dashboard_Xbox_Game_Pass.xlsx.
- Habilite macros ao abrir o arquivo.
- Use os filtros para interagir com os dados ou clique em “Limpar Filtros” para restaurar.
- Utilize o botão de Tela Cheia para uma visualização mais limpa.

### 📌 Observações
- Projeto feito com fins educacionais.
- Os dados utilizados são fictícios, representando um cenário simulado de vendas da Microsoft XBOX.

### 👩‍💻 Autor(a)
Projeto criado por Alais Cassimira Salino Barbosa como parte do desafio de dashboard no Excel do Bootcamp da DIO - Santander - Excel com Inteligência Artificial em junho de 2025.
