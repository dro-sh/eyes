Dim CurDir
Dim bb

Sub SetGlobals(aCurDir, aBB)
  CurDir = aCurDir
  Set bb = aBB
End Sub 


Sub ShowBB()
  bb.ShowObjectTree()
End Sub 

Sub Form(name, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20, t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31, t32, t33, t34, t35, t36, t37, t38, t39, t40, t41, t42, t43, t44, pr1, pr2, pr3, pr4, pr5, as1, as2, as3, as4, as5, as6, as7, as8, as9, as10, as11, as12, bs1, bs2, bs3, bs4, bs5, bs6, bs7, bs8, bs9, bs10, bs11, bs12, bs13, ai1, ai2, ai3, ai4, ai5, bi1, bi2, bi3, bi4, i1, i2, norm)
  Dim App
  Set App = CreateObject("Excel.Application")  
  App.Workbooks.Open(CurDir + "rep.xls")

  Dim symptomes : Set symptomes = CreateObject("System.Collections.ArrayList")
  Dim addProcedures : Set addProcedures = CreateObject("System.Collections.ArrayList")
  Dim illnesses : Set illnesses = CreateObject("System.Collections.ArrayList")
  Dim treatments : Set treatments = CreateObject("System.Collections.ArrayList")

  if (s1 = "да") then
    symptomes.Add "выраженное покраснение ГЯ"
  end if
  if (s2 = "да") then
    symptomes.Add "непрекращающийся зуд"
  end if
  if (s3 = "да") then
    symptomes.Add "синдром песка в глазах"
  end if
  if (s4 = "да") then
    symptomes.Add "изображение не в фокусе"
  end if
  if (s5 = "да") then
    symptomes.Add "боль при напряжении глазных мышц"
  end if
  if (s6 = "да") then
    symptomes.Add "головная боль"
  end if
  if (s7 = "да") then
    symptomes.Add "нагноение в области слезного мешка ГЯ"
  end if
  if (s8 = "да") then
    symptomes.Add "признаки аллергической реакции"
  end if
  if (s9 = "да") then
    symptomes.Add "повышенная температура"
  end if
  if (s10 = "да") then
    symptomes.Add "светобоязнь"
  end if
  if (s11 = "да") then
    symptomes.Add "пациент не может рассмотреть объект на близком расстоянии"
  end if
  if (s12 = "да") then
    symptomes.Add "близко расположенные предметы расплываются"
  end if
  if (s13 = "да") then
    symptomes.Add "контур предметов становится размытым"
  end if
  if (s14 = "да") then
    symptomes.Add "есть ощущение внутреннего давления"
  end if
  if (s15 = "да") then
    symptomes.Add "при взгляде на источник света появляются радужные круги"
  end if
  if (s16 = "да") then
    symptomes.Add "выпадают ресницы"
  end if
  if (s17 = "да") then
    symptomes.Add "покраснение и опухание век"
  end if
  if (s18 = "да") then
    symptomes.Add "при визуальном осмотре на ресницах и веках видны гнойные чешуйки"
  end if
  if (s19 = "да") then
    symptomes.Add "положительный анализ на хламидии"
  end if
  if (s20 = "да") then
    symptomes.Add "повышены референсные значения лейкоцитов"
  end if
  if (as1 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А на месте локализации видно уплотнения структуры"
  end if
  if (as2 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А ультразвуковая волна направлена параллельно зрительной оси"
  end if
  if (as3 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А поверхность макула гладкая"
  end if
  if (as4 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А эхосигнал возвращается к излучателю, где пик высокой амплитуды"
  end if
  if (as5 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А поверхность макулы выпуклая"
  end if
  if (as6 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А часть эха отражается далеко от излучателя, где пик более низкой амплитуды"
  end if
  if (as7 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А поверхность макулы нерегулярная"
  end if
  if (as8 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А часть эха отражается далеко от излучателя, где пик низкой амплитуды"
  end if
  if (as9 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А ближайшая точка p.p отодвигается от глаза"
  end if
  if (as10 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А часть эха отражается далеко от излучателя, где пик низкой амплитуды"
  end if
  if (as11 = "да") then
    symptomes.Add "на УЗ-изображении в режиме А уменьшается общий объем аккомодации"
  end if
  if (as12 = "да") then
    symptomes.Add "показатели оттока ВГЖ выше установленных норм"
  end if
  if (bs1 = "да") then
    symptomes.Add "выраженное покраснение ГЯ"
  end if
  if (bs2 = "да") then
    symptomes.Add "непрекращающийся зуд"
  end if
  if (bs3 = "да") then
    symptomes.Add "синдром песка в глазах"
  end if
  if (bs4 = "да") then
    symptomes.Add "изображение не в фокусе"
  end if
  if (bs5 = "да") then
    symptomes.Add "боль при напряжении глазных мышц"
  end if
  if (bs6 = "да") then
    symptomes.Add "головная боль"
  end if
  if (bs7 = "да") then
    symptomes.Add "нагноение в области слезного мешкаГЯ"
  end if
  if (bs8 = "да") then
    symptomes.Add "признаки аллергической реакции"
  end if
  if (bs9 = "да") then
    symptomes.Add "повышенная температура"
  end if
  if (bs10 = "да") then
    symptomes.Add "светобоязнь"
  end if
  if (bs11 = "да") then
    symptomes.Add "пациент не может рассмотреть объект на близком расстоянии"
  end if
  if (bs12 = "да") then
    symptomes.Add "близко расположенные предметы расплываются"
  end if
  if (bs13 = "да") then
    symptomes.Add "контур предметов становится размытым"
  end if

  if (pr1 = "да") then
    addProcedures.Add "провести УЗИ ГЯ в режиме А"
  end if
  if (pr2 = "да") then
    addProcedures.Add "сдать анализ на хламидии"
  end if
  if (pr3 = "да") then
    addProcedures.Add "сдать общий анализ крови"
  end if
  if (pr4 = "да") then
    addProcedures.Add "исследовать показатели оттока ВГЖ с помощью тонографа"
  end if
  if (pr5 = "да") then
    addProcedures.Add "провести УЗИ ГЯ в режиме Б"
  end if

  if (tr1 = "да") then
    treatments.Add "промыть ГЯ проточной водой"
  end if
  if (t2 = "да") then
    treatments.Add "осмотреть ГЯ на наличие инородных предметов"
  end if
  if (t3 = "да") then
    treatments.Add "выполнить гимнастику для глаз"
  end if
  if (t4 = "да") then
    treatments.Add "выписать капли артикаин"
  end if
  if (t5 = "да") then
    treatments.Add "выписать капли 2% лидокоин"
  end if
  if (t6 = "да") then
    treatments.Add "снять напряжение путем наложения светонепроницаемой ткани"
  end if
  if (t7 = "да") then
    treatments.Add "промыть ГЯ кипяченной водой"
  end if
  if (t8 = "да") then
    treatments.Add "выписать капли лавомицитин"
  end if
  if (t9 = "да") then
    treatments.Add "выписать капли ципромет"
  end if
  if (t10 = "да") then
    treatments.Add "наложить тетрациклиновую глазную мазь"
  end if
  if (t11 = "да") then
    treatments.Add "наложить эритромициновую глазную мазь"
  end if
  if (t12 = "да") then
    treatments.Add "выписать левофлюксацин"
  end if
  if (t13 = "да") then
    treatments.Add "провести физикальный осмотр с использованием щелевой лампы"
  end if
  if (t14 = "да") then
    treatments.Add "провести физикальный осмотр с использованием офтальмоскопа"
  end if
  if (t15 = "да") then
    treatments.Add "оценить передней камеры с помощью гониоскопии"
  end if
  if (t16 = "да") then
    treatments.Add "проверить поля зрения"
  end if
  if (t17 = "да") then
    treatments.Add "исследовать показатели оттока ВГЖ с помощью тонографа"
  end if
  if (t19 = "да") then
    treatments.Add "промыть слезные пути"
  end if
  if (t20 = "да") then
    treatments.Add "визуальный осмотр"
  end if
  if (t21 = "да") then
    treatments.Add "произвести бакпосев методом ПЦР"
  end if
  if (t22 = "да") then
    treatments.Add "произвести бакпосев методом ИФА"
  end if
  if (t23 = "да") then
    treatments.Add "выписать рецепт на приобретение очков"
  end if
  if (t24 = "да") then
    treatments.Add "выписать тобрекс"
  end if
  if (t25 = "да") then
    treatments.Add "рекомендуется потребление комплекса витаминов (A, B1, B6, E)"
  end if
  if (t26 = "да") then
    treatments.Add "рекомендуется соблюдение рациона питания, содержащего необходимые микроэлементы, правил гигиены, гимнастика для глаз при долгом чтении / работе за компьютером"
  end if
  if (t27 = "да") then
    treatments.Add "показано хирургическое лечение"
  end if
  if (t28 = "да") then
    treatments.Add "лучевая терапия + клеточная терапия"
  end if
  if (t29 = "да") then
    illnesses.Add "первичная отслойка"
  end if
  if (t30 = "да") then
    treatments.Add "фотолазеркриодиатермокоагуляция краев и зоны разрыва (отрыва) сетчатки"
  end if
  if (t31 = "да") then
    treatments.Add "склеропластические операции (рифление, наложение кругового шва, пломбирование и др.)"
  end if
  if (t32 = "да") then
    treatments.Add "консультация невролога"
  end if
  if (t33 = "да") then
    treatments.Add "снижение ВЧД"
  end if
  if (t34 = "да") then
    treatments.Add "стероидные гармоны"
  end if
  if (t35 = "да") then
    treatments.Add "нестероидные противовоспалительные препараты"
  end if
  if (t36 = "да") then
    treatments.Add "выписать квинакс"
  end if
  if (t37 = "да") then
    treatments.Add "выписать визотимин"
  end if
  if (t38 = "да") then
    treatments.Add "необходима факоэмульсификация"
  end if
  if (t39 = "да") then
    treatments.Add "выписать ципрофлоксацина"
  end if
  if (t40 = "да") then
    treatments.Add "выписать эритромициновую мазь"
  end if
  if (t41 = "да") then
    treatments.Add "выписать капли левомицитин"
  end if
  if (t42 = "да") then
    treatments.Add "выписать капли ципромет"
  end if
  if (t43 = "да") then
    treatments.Add "выписать тетрациклиновую глазную мазь"
  end if
  if (t44 = "да") then
    treatments.Add "выписать эритромициновую глазную мазь"
  end if

  if (ai1 = "да") then
    illnesses.Add "гиперметропия"
  end if
  if (ai2 = "да") then
    illnesses.Add "макулярный отек"
  end if
  if (ai3 = "да") then
    illnesses.Add "отслойка пигментного эпителия сетчатки"
  end if
  if (ai4 = "да") then
    illnesses.Add "макулярная эпиретинальная мембрана"
  end if
  if (ai5 = "да") then
    illnesses.Add "выраженная миопия"
  end if

  if (bi1 = "да") then
    illnesses.Add "ретинобластома"
  end if
  if (bi2 = "да") then
    illnesses.Add "вторичная отслойка"
  end if
  if (bi3 = "да") then
    illnesses.Add "фиброз"
  end if
  if (bi4 = "да") then
    illnesses.Add "катаракта"
  end if

  if (i1 = "да") then
    illnesses.Add "блефарит"
  end if
  if (i2 = "да") then
    illnesses.Add "хламидиозный конъюнктивит"
  end if

  if ((norm = "да") or (symptomes.Count = 0)) then
    illnesses.Add "норма"
  end if

  App.Range("B2").Select
  App.ActiveCell.Value = name
  App.Range("C2").Select
  App.ActiveCell.Value = CStr(Date())
  
  Dim idx 
  idx = 4
  Dim s
  For Each s In symptomes
    App.Range("A" & CStr(idx)).Select
    App.ActiveCell.Value = s

    idx = idx + 1
  Next

  idx = 4
  Dim p
  For Each p In addProcedures
    App.Range("B" & CStr(idx)).Select
    App.ActiveCell.Value = p

    idx = idx + 1
  Next 
  
  idx = 4
  Dim ill
  For Each ill In illnesses
    App.Range("C" & CStr(idx)).Select
    App.ActiveCell.Value = ill

    idx = idx + 1
  Next

  idx = 4
  Dim tr

  For Each tr In treatments
    App.Range("D" & CStr(idx)).Select
    App.ActiveCell.Value = tr

    idx = idx + 1
  Next
  
  App.Range("A1").Select
  App.Visible = true
  App = ""
  'App.Quit
End Sub

