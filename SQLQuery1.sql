select A.*,B.company,B.class from (
select name , count(name) as cnt from mapresult group by name ) A
right join mapresult B on A.name=B.name order by A.cnt desc

select A.*,B.name,B.class from (
select company ,count(company) as cnt  from mapresult group by company) A
right join mapresult B on A.company=B.company order by A.cnt desc,A.company


select A.*,B.name,B.class from (
select company   from mapresult group by company) A
right join mapresult B on A.company=B.company order by B.class desc,A.company 



select company ,count(company) as cnt  from mapresult group by company order by count(company) desc

--delete mapresult
