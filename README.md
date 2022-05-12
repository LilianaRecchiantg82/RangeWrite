# RangeWrite
Local $sFormula = _Excel_RangeRead($oWorkBook, $oWorkSheet, "A1", 2) ; Note: 2 means get the formula _Excel_RangeWrite($oWorkBook, $oWorkSheet, $sFormula, "A1")
