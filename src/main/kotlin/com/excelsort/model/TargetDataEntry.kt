package com.excelsort.model

import java.math.BigDecimal

data class TargetDataEntry(
    val nazwaProduktu: String,
    val ilosc: BigDecimal?,
    val dostepnosc: BigDecimal?,
    val ean: BigDecimal?
)
