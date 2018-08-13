create or replace package tq84 as
    function func(
               p_one varchar2 := 'default',
               p_two number   := 42
            )
            return varchar2;
end tq84;
/

create or replace package body tq84 as

  function func(p_one varchar2 := 'default', p_two number := 42) return varchar2 is
  begin
      return 'p_one = ' || p_one || ', p_two = ' || p_two;
  end func;

end tq84;
/
