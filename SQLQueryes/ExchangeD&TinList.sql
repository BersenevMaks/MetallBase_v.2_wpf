select * from Product where [Name]='����' and Diametr>Tolshina and Diametr>500 and Tolshina>0
update Product
set Diametr = Tolshina, Tolshina = Diametr
where [Name]='����' and Diametr>Tolshina and Diametr>100 and Tolshina>0
select * from Product where [Name]='����'