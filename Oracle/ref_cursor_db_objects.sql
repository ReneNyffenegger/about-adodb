create table refcursor_test_tbl (
  site_id         number(2)       not null,
  location        varchar2(12)
);


insert into refcursor_test_tbl values (1, 'Paris'     );
insert into refcursor_test_tbl values (2, 'Boston'    );
insert into refcursor_test_tbl values (3, 'London'    );
insert into refcursor_test_tbl values (4, 'Stockholm' );
insert into refcursor_test_tbl values (5, 'Ottawa'    );
insert into refcursor_test_tbl values (6, 'Washington');
insert into refcursor_test_tbl values (7, 'La'        );
insert into refcursor_test_tbl values (8, 'Toronto'   );
insert into refcursor_test_tbl values (3, 'Zuerich'   );


create or replace package tq84_refcursor_test_pck as -- {
     
  procedure proc_1 (
          Pmyid      in number,
          Pmycursor  out sys_refcursor, -- use cursor
          x          out sys_refcursor,
          Perrorcode out number);    
         
  procedure proc_2 (
          Pmyquery   in  varchar2,
          Pmycursor  out sys_refcursor,
          Perrorcode out number);

end tq84_refcursor_test_pck; -- }
/


create or replace package body tq84_refcursor_test_pck as -- {

   procedure proc_1 (
       Pmyid      in number,
       Pmycursor  out sys_refcursor, -- use cursor
       x          out sys_refcursor, -- use cursor
       Perrorcode out number) is
   begin
          Perrorcode := 0;
          -- Open the ref cursor
          -- Use Input Variable "PmyID" as part of the query.

     open pmycursor for
          select 'foo' foo, location
          from refcursor_test_tbl
                where site_id = pmyid;

     open x for select * from refcursor_test_tbl where site_id < 5;
            
   exception when others then
       Perrorcode := SQLCODE;  
   end proc_1; 
    
   procedure proc_2 (
        Pmyquery in varchar2,
        Pmycursor out sys_refcursor,
        Perrorcode out number) 
   is 
   begin
         Perrorcode := 0;
         -- Open the REF CURSOR 
         -- This procedure uses a query
         -- which is passed in as a parameter.
  
         open pmycursor for pmyquery;
    
   exception when others then
           Perrorcode := sqlcode;
   end proc_2;

end tq84_refcursor_test_pck; -- }
/

