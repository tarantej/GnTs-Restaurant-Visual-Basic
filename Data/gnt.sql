create database gnt
use gnt
create table gnt_login
(
username varchar(10),
password varchar(10),
role varchar(10) 
);
alter table gnt_login drop column role
select * from login
create table reserve
(
res_date datetime,
cust_name varchar(20),
table_for int,
res_time time,
res_status varchar(5)
);
select * from reserve
create table order_taking
(
item_code varchar(10),
item_name varchar(20),
item_quantity int
);
select* from order_taking
create table bill
(
billID varchar (10),
billDate datetime,
itemcode varchar(10),
itemname varchar(20),
itemquantity int,
itemprice int,
totalamt int,
vat int,
totalbill int
)
select * from bill
create table item
(
ItemID varchar(10),
ItemName varchar(30),
ItemQuantity int,
ItemPrice int,
ItemStatus varchar(10)
)
select * from item

create table emp_detail
(
Emp_ID varchar(10),
Emp_Name varchar(30),
Emp_Role varchar(10),
Emp_Address varchar(50),
Emp_City varchar(50),
Emp_CityPin int,
Emp_Mobile int,
Emp_ResNum int,
Emp_Join date,
Emp_Salary int
)
select * from emp_detail 
create table new_emp
(
Emp_Name varchar(30),
Emp_Gender char(10),
Emp_Age int,
Emp_Married char(5),
Emp_Qual varchar(100),
Emp_Address varchar(50),
Emp_City varchar(50),
Emp_CityPin int,
Emp_Mobile int,
Emp_ResNum int,
Emp_Join date
)
select * from new_emp 
create table emp_entry
(
currdate date,
Emp_ID varchar(10),
Emp_Name varchar(30),
Emp_Role varchar(10),
TimeIn time,
TimeOut time
)
select * from emp_entry 

create table product_status
(
prod_id varchar (10),
prod_name varchar(50),
prod_status varchar(30)
)
create table material_stock
(
mat_id varchar(10),
mat_name varchar(30),
mat_type varchar(10),
mat_stock int,
mat_stock_qty varchar(50)
)
create table overall_stock
(
item_name varchar(50),
stock int,
quantity_type varchar(50)
)

insert into gnt_login values ('tarantej','singh')
insert into gnt_login values('mona','sharma')

select * from gnt_login

