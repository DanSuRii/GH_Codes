    #"Rounded Off" = Table.TransformColumns(#"Changed Type1",{}),
    #"Rounded Off1" = Table.TransformColumns(#"Added Conditional Column",{{"Custom", each Number.Round(_, 2), type number}, {"Diff_Lei", each Number.Round(_, 2), Currency.Type}, {"Diff_Absat", each Number.Round(_, 2), Currency.Type}})
