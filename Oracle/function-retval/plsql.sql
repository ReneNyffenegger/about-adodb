create or replace package tq84 as
    function func return varchar2;
end tq84;
/

create or replace package body tq84 as
    function func return varchar2 is
  begin
      return 'Hello world!';
  end func;
end tq84;
/
