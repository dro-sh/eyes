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

  if (s1 = "��") then
    symptomes.Add "���������� ����������� ��"
  end if
  if (s2 = "��") then
    symptomes.Add "���������������� ���"
  end if
  if (s3 = "��") then
    symptomes.Add "������� ����� � ������"
  end if
  if (s4 = "��") then
    symptomes.Add "����������� �� � ������"
  end if
  if (s5 = "��") then
    symptomes.Add "���� ��� ���������� ������� ����"
  end if
  if (s6 = "��") then
    symptomes.Add "�������� ����"
  end if
  if (s7 = "��") then
    symptomes.Add "��������� � ������� �������� ����� ��"
  end if
  if (s8 = "��") then
    symptomes.Add "�������� ������������� �������"
  end if
  if (s9 = "��") then
    symptomes.Add "���������� �����������"
  end if
  if (s10 = "��") then
    symptomes.Add "�����������"
  end if
  if (s11 = "��") then
    symptomes.Add "������� �� ����� ����������� ������ �� ������� ����������"
  end if
  if (s12 = "��") then
    symptomes.Add "������ ������������� �������� ������������"
  end if
  if (s13 = "��") then
    symptomes.Add "������ ��������� ���������� ��������"
  end if
  if (s14 = "��") then
    symptomes.Add "���� �������� ����������� ��������"
  end if
  if (s15 = "��") then
    symptomes.Add "��� ������� �� �������� ����� ���������� �������� �����"
  end if
  if (s16 = "��") then
    symptomes.Add "�������� �������"
  end if
  if (s17 = "��") then
    symptomes.Add "����������� � �������� ���"
  end if
  if (s18 = "��") then
    symptomes.Add "��� ���������� ������� �� �������� � ����� ����� ������� �������"
  end if
  if (s19 = "��") then
    symptomes.Add "������������� ������ �� ��������"
  end if
  if (s20 = "��") then
    symptomes.Add "�������� ����������� �������� ����������"
  end if
  if (as1 = "��") then
    symptomes.Add "�� ��-����������� � ������ � �� ����� ����������� ����� ���������� ���������"
  end if
  if (as2 = "��") then
    symptomes.Add "�� ��-����������� � ������ � �������������� ����� ���������� ����������� ���������� ���"
  end if
  if (as3 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����������� ������ �������"
  end if
  if (as4 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ��������� ������������ � ����������, ��� ��� ������� ���������"
  end if
  if (as5 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����������� ������ ��������"
  end if
  if (as6 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����� ��� ���������� ������ �� ����������, ��� ��� ����� ������ ���������"
  end if
  if (as7 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����������� ������ ������������"
  end if
  if (as8 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����� ��� ���������� ������ �� ����������, ��� ��� ������ ���������"
  end if
  if (as9 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ��������� ����� p.p ������������ �� �����"
  end if
  if (as10 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����� ��� ���������� ������ �� ����������, ��� ��� ������ ���������"
  end if
  if (as11 = "��") then
    symptomes.Add "�� ��-����������� � ������ � ����������� ����� ����� �����������"
  end if
  if (as12 = "��") then
    symptomes.Add "���������� ������ ��� ���� ������������� ����"
  end if
  if (bs1 = "��") then
    symptomes.Add "���������� ����������� ��"
  end if
  if (bs2 = "��") then
    symptomes.Add "���������������� ���"
  end if
  if (bs3 = "��") then
    symptomes.Add "������� ����� � ������"
  end if
  if (bs4 = "��") then
    symptomes.Add "����������� �� � ������"
  end if
  if (bs5 = "��") then
    symptomes.Add "���� ��� ���������� ������� ����"
  end if
  if (bs6 = "��") then
    symptomes.Add "�������� ����"
  end if
  if (bs7 = "��") then
    symptomes.Add "��������� � ������� �������� �������"
  end if
  if (bs8 = "��") then
    symptomes.Add "�������� ������������� �������"
  end if
  if (bs9 = "��") then
    symptomes.Add "���������� �����������"
  end if
  if (bs10 = "��") then
    symptomes.Add "�����������"
  end if
  if (bs11 = "��") then
    symptomes.Add "������� �� ����� ����������� ������ �� ������� ����������"
  end if
  if (bs12 = "��") then
    symptomes.Add "������ ������������� �������� ������������"
  end if
  if (bs13 = "��") then
    symptomes.Add "������ ��������� ���������� ��������"
  end if

  if (pr1 = "��") then
    addProcedures.Add "�������� ��� �� � ������ �"
  end if
  if (pr2 = "��") then
    addProcedures.Add "����� ������ �� ��������"
  end if
  if (pr3 = "��") then
    addProcedures.Add "����� ����� ������ �����"
  end if
  if (pr4 = "��") then
    addProcedures.Add "����������� ���������� ������ ��� � ������� ���������"
  end if
  if (pr5 = "��") then
    addProcedures.Add "�������� ��� �� � ������ �"
  end if

  if (tr1 = "��") then
    treatments.Add "������� �� ��������� �����"
  end if
  if (t2 = "��") then
    treatments.Add "��������� �� �� ������� ��������� ���������"
  end if
  if (t3 = "��") then
    treatments.Add "��������� ���������� ��� ����"
  end if
  if (t4 = "��") then
    treatments.Add "�������� ����� ��������"
  end if
  if (t5 = "��") then
    treatments.Add "�������� ����� 2% ��������"
  end if
  if (t6 = "��") then
    treatments.Add "����� ���������� ����� ��������� ������������������ �����"
  end if
  if (t7 = "��") then
    treatments.Add "������� �� ���������� �����"
  end if
  if (t8 = "��") then
    treatments.Add "�������� ����� �����������"
  end if
  if (t9 = "��") then
    treatments.Add "�������� ����� ��������"
  end if
  if (t10 = "��") then
    treatments.Add "�������� ��������������� ������� ����"
  end if
  if (t11 = "��") then
    treatments.Add "�������� ��������������� ������� ����"
  end if
  if (t12 = "��") then
    treatments.Add "�������� �������������"
  end if
  if (t13 = "��") then
    treatments.Add "�������� ����������� ������ � �������������� ������� �����"
  end if
  if (t14 = "��") then
    treatments.Add "�������� ����������� ������ � �������������� �������������"
  end if
  if (t15 = "��") then
    treatments.Add "������� �������� ������ � ������� �����������"
  end if
  if (t16 = "��") then
    treatments.Add "��������� ���� ������"
  end if
  if (t17 = "��") then
    treatments.Add "����������� ���������� ������ ��� � ������� ���������"
  end if
  if (t19 = "��") then
    treatments.Add "������� ������� ����"
  end if
  if (t20 = "��") then
    treatments.Add "���������� ������"
  end if
  if (t21 = "��") then
    treatments.Add "���������� �������� ������� ���"
  end if
  if (t22 = "��") then
    treatments.Add "���������� �������� ������� ���"
  end if
  if (t23 = "��") then
    treatments.Add "�������� ������ �� ������������ �����"
  end if
  if (t24 = "��") then
    treatments.Add "�������� �������"
  end if
  if (t25 = "��") then
    treatments.Add "������������� ����������� ��������� ��������� (A, B1, B6, E)"
  end if
  if (t26 = "��") then
    treatments.Add "������������� ���������� ������� �������, ����������� ����������� �������������, ������ �������, ���������� ��� ���� ��� ������ ������ / ������ �� �����������"
  end if
  if (t27 = "��") then
    treatments.Add "�������� ������������� �������"
  end if
  if (t28 = "��") then
    treatments.Add "������� ������� + ��������� �������"
  end if
  if (t29 = "��") then
    illnesses.Add "��������� ��������"
  end if
  if (t30 = "��") then
    treatments.Add "������������������������������� ����� � ���� ������� (������) ��������"
  end if
  if (t31 = "��") then
    treatments.Add "������������������ �������� (��������, ��������� ��������� ���, ������������� � ��.)"
  end if
  if (t32 = "��") then
    treatments.Add "������������ ���������"
  end if
  if (t33 = "��") then
    treatments.Add "�������� ���"
  end if
  if (t34 = "��") then
    treatments.Add "���������� �������"
  end if
  if (t35 = "��") then
    treatments.Add "������������ ��������������������� ���������"
  end if
  if (t36 = "��") then
    treatments.Add "�������� �������"
  end if
  if (t37 = "��") then
    treatments.Add "�������� ���������"
  end if
  if (t38 = "��") then
    treatments.Add "���������� ������������������"
  end if
  if (t39 = "��") then
    treatments.Add "�������� ���������������"
  end if
  if (t40 = "��") then
    treatments.Add "�������� ��������������� ����"
  end if
  if (t41 = "��") then
    treatments.Add "�������� ����� �����������"
  end if
  if (t42 = "��") then
    treatments.Add "�������� ����� ��������"
  end if
  if (t43 = "��") then
    treatments.Add "�������� ��������������� ������� ����"
  end if
  if (t44 = "��") then
    treatments.Add "�������� ��������������� ������� ����"
  end if

  if (ai1 = "��") then
    illnesses.Add "�������������"
  end if
  if (ai2 = "��") then
    illnesses.Add "���������� ����"
  end if
  if (ai3 = "��") then
    illnesses.Add "�������� ����������� �������� ��������"
  end if
  if (ai4 = "��") then
    illnesses.Add "���������� �������������� ��������"
  end if
  if (ai5 = "��") then
    illnesses.Add "���������� ������"
  end if

  if (bi1 = "��") then
    illnesses.Add "��������������"
  end if
  if (bi2 = "��") then
    illnesses.Add "��������� ��������"
  end if
  if (bi3 = "��") then
    illnesses.Add "������"
  end if
  if (bi4 = "��") then
    illnesses.Add "���������"
  end if

  if (i1 = "��") then
    illnesses.Add "��������"
  end if
  if (i2 = "��") then
    illnesses.Add "������������ ������������"
  end if

  if ((norm = "��") or (symptomes.Count = 0)) then
    illnesses.Add "�����"
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

