drop table     tq84_table;
drop procedure tq84_procedure;

create table tq84_table (a number);

create procedure tq84_procedure(a in number, b in number) as
begin

  for i in  a .. b loop
    insert into tq84_table values (i);
  end loop;

end;
/
