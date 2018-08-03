create package tq84_in_out_parameters as
       procedure proc(
          param_one   in  number,
          param_two   out number,
          param_three out number
       );
end tq84_in_out_parameters;
/

create package body tq84_in_out_parameters as
       procedure proc(
          param_one   in  number,
          param_two   out number,
          param_three out number
       ) is
       begin

          param_two   := param_one * 2;
          param_three := param_one * 3;

       end proc;
end tq84_in_out_parameters;
/
