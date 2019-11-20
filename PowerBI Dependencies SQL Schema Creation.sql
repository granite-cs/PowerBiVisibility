/*
(c) 2019 David Berglin 
This file is part of the PowerBiVisibility project.
PowerBiVisibility is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
PowerBiVisibility is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with PowerBiVisibility.  If not, see https://www.gnu.org/licenses/.
*/

IF OBJECT_ID('[dbo].[Dependencies]', 'U') IS NULL -- U = Table... V = View... P = Stored Procedure
begin
    print 'creating table [dbo].[Dependencies]' + ' ...' + convert(varchar, getdate(), 121)

    CREATE TABLE [dbo].[Dependencies](
	    [DependencyId]      [int]       IDENTITY(1,1)   NOT NULL,
	    [Source]            [varchar](400)              NOT NULL,   -- Where did this dependency come from (which PowerBI file or SSAS server ?)
	    [ParentLocation]    [varchar](400)              NOT NULL,   -- What tab(PowerBI)    or table(SSAS)          holds the dependency?
	    [ParentName]        [varchar](400)              NOT NULL,   -- What visual(PowerBI) or measure/column(SSAS) holds the dependency?
	    [ParentAddress]     [varchar](400)              NOT NULL,   -- Standardized 'Location'[Name]
	    [ParentType]        [varchar](40)               NOT NULL,   -- column/measure/visual
	    [ChildLocation]     [varchar](400)              NULL,       -- What table(SSAS)          is depended upon by the parent?
	    [ChildName]         [varchar](400)              NULL,       -- What measure/column(SSAS) is depended upon by the parent?
	    [ChildAddress]      [varchar](400)              NULL,       -- Standardized 'Location'[Name]
	    [ChildType]         [varchar](40)               NULL,       -- column/measure
	    [Content]           [nvarchar](4000)            NULL,       -- the plan is to add Measure code into this column
     CONSTRAINT [PK_Dependencies] PRIMARY KEY CLUSTERED 
    (
	    [DependencyId] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
end
GO

IF OBJECT_ID('[dbo].[vDependency_Scripts]', 'V') IS NOT NULL -- U = Table... V = View... P = Stored Procedure
begin
    print 'deleting view [dbo].[vDependency_Scripts]' + ' ...' + convert(varchar, getdate(), 121)
    drop view [dbo].[vDependency_Scripts]
end
GO
    print 'creating view [dbo].[vDependency_Scripts]' + ' ...' + convert(varchar, getdate(), 121)
GO
CREATE VIEW [dbo].[vDependency_Scripts] AS
/*  
    (c) 2019 David Berglin 
    This file is part of the PowerBiVisibility project.
    PowerBiVisibility is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
    PowerBiVisibility is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
    You should have received a copy of the GNU General Public License along with PowerBiVisibility.  If not, see https://www.gnu.org/licenses/.
    

    Join to this view and replace values as needed, to generate scripts about dependencies 


    usage example:

    select 
        *
        , Replace(Replace(s.Scripts, '[ReplaceWithName]', DependencyChild), '[ReplaceWithScriptableName]', ScriptableDependencyChild)   as [Child Scripts]
        , Replace(Replace(s.Scripts, '[ReplaceWithName]', ParentAddress),   '[ReplaceWithScriptableName]', ScriptableParent)            as [Parent Scripts]
    from cte_nested c 
    join vDependency_Scripts s on 1 = 1

*/
select '
-- [ReplaceWithName] ...scripts below 

-- ALL my Child Dependencies
Select * 
from [dbo].[vDependency_Children] 
where TopMostParent   = ''[ReplaceWithScriptableName]''
order by ChildAddress

-- All Parents Who Depend on me
Select * 
from [dbo].[vDependency_Parents]  
where DependencyChild = ''[ReplaceWithScriptableName]''
order by ParentAddress

-- ALL my DISTINCT Child Dependencies
Select DISTINCT 
     TopMostParent
    ,ChildType
    ,ChildAddress 
from [dbo].[vDependency_Children] 
where TopMostParent   = ''[ReplaceWithScriptableName]''
order by ChildAddress

-- All DISTINCT Parents Who Depend on me
Select DISTINCT 
     DependencyChild
    ,ParentType
    ,ParentAddress 
from [dbo].[vDependency_Parents]  
where DependencyChild = ''[ReplaceWithScriptableName]''
order by ParentAddress

-- Raw Depencency Data about me
Select * 
from [dbo].[Dependencies]
where ChildAddress  = ''[ReplaceWithScriptableName]''
   or ParentAddress = ''[ReplaceWithScriptableName]''
Order By case when ChildAddress  = ''[ReplaceWithScriptableName]'' then 0 else 1 end -- put parents at top, children at bottom
        ' as [Scripts]

GO

IF OBJECT_ID('[dbo].[vDependency_Parents]', 'V') IS NOT NULL -- U = Table... V = View... P = Stored Procedure
begin
    print 'deleting view [dbo].[vDependency_Parents]' + ' ...' + convert(varchar, getdate(), 121)
    drop view [dbo].[vDependency_Parents]
end
GO
    print 'creating view [dbo].[vDependency_Parents]' + ' ...' + convert(varchar, getdate(), 121)
GO
CREATE VIEW [dbo].[vDependency_Parents] AS
/*
--------------------------------------------------------------
(c) 2019 David Berglin 
This file is part of the PowerBiVisibility project.
PowerBiVisibility is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
PowerBiVisibility is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with PowerBiVisibility.  If not, see https://www.gnu.org/licenses/.
--------------------------------------------------------------
Purpose is to Find Visuals/Measures (who are Parents) by Dependencies (Children)
        This view will show PowerBI Visuals and SSAS Measures which depend on a particular item. 

 populate [dbo].[Dependencies] using this Powershell scripts: 
    PowerBI Visibility.ps1        ... use SQL output option

 research...
    select * from [dbo].[Dependencies] 
    select * from [dbo].[Dependencies] where ChildAddress like '%[[]Total Hours]%'  -- [[]text] handles square bracket in "Like" clause

 Sample Script
    select distinct *
        ,'Select * from [dbo].[vDependency_Children] where TopMostParent   = '''+Replace(ParentAddress, '''', '''''')+'''' as [Script To Find Children Dependencies]
        ,'Select * from [dbo].[vDependency_Parents]  where DependencyChild = '''+Replace(ParentAddress, '''', '''''')+'''' as [Script To Find Parents Who Depend on Me]
    from [dbo].[vDependency_Parents]
    order by ParentAddress

Distinct Script
    Select distinct
         DependencyChild
        ,ParentType
        ,ParentAddress
    from [dbo].[vDependency_Parents]
--------------------------------------------------------------
*/

with cte_nested as ( -- holds nested dependency objects
    
    -- Select the 'normal' (not recursive, just top level) rows of data
    select 
         1 as Recursion
        ,d.ChildAddress as DependencyChild
        ,Replace(d.ParentAddress, '''', '''''') as ScriptableParent -- replace single quotes with two single quotes... for pasting into SQL Select Scripts
        ,Replace(d.ChildAddress, '''', '''''')  as ScriptableDependencyChild -- replace single quotes with two single quotes... for pasting into SQL Select Scripts
        ,d.ChildType 
        ,d.ChildAddress
        ,d.ChildLocation
        ,d.ChildName
        ,d.ParentType    -- parent
        ,d.ParentAddress -- parent
        ,convert(varchar(5000), d.ParentAddress + isnull('
<= ' + d.ChildAddress,'')) as [Trail]
    from [dbo].[Dependencies] d 
    where d.ChildAddress is not null -- no need to pull in null dependencies
        and d.ChildAddress <> ''
    
    union all 

    -- Recursively reach back into this same CTE to find nested data
    select 
         n.Recursion + 1
        ,n.DependencyChild
        ,Replace(d.ParentAddress, '''', '''''') as ScriptableParent
        ,n.ScriptableDependencyChild
        ,d.ChildType 
        ,d.ChildAddress
        ,d.ChildLocation
        ,d.ChildName
        ,d.ParentType    -- parent
        ,d.ParentAddress -- parent
        ,convert(varchar(5000), d.ParentAddress + isnull(' 
<= ' + n.[Trail], '')) 
    from [dbo].[Dependencies] d
    join cte_nested n
        on  n.ParentAddress     =   d.ChildAddress -- pull in the parent
        and d.ParentAddress     <> ''
        and n.Recursion         <   20

)

----------------------------------------------------------------
-- Get final results
----------------------------------------------------------------
    select distinct    
         c.DependencyChild
        ,c.ParentType
        ,c.ParentAddress
        ,c.Recursion
        ,c.[Trail] + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) as [Trail]
        , Replace(Replace(s.Scripts, '[ReplaceWithName]', DependencyChild), '[ReplaceWithScriptableName]', ScriptableDependencyChild)   as [Child Scripts]
        , Replace(Replace(s.Scripts, '[ReplaceWithName]', ParentAddress),   '[ReplaceWithScriptableName]', ScriptableParent)            as [Parent Scripts]
    from cte_nested c 
    join vDependency_Scripts s on 1 = 1 -- no real join needed, this copies the simple text of the script generator, and replaces values to produce scripts you can copy/paste


GO
    print 'finished view [dbo].[vDependency_Parents]' + ' ...' + convert(varchar, getdate(), 121)


GO
IF OBJECT_ID('[dbo].[vDependency_Children]', 'V') IS NOT NULL -- U = Table... V = View... P = Stored Procedure
begin
    print 'deleting view [dbo].[vDependency_Children]' + ' ...' + convert(varchar, getdate(), 121)
    drop view [dbo].[vDependency_Children]
end
GO
    print 'creating view [dbo].[vDependency_Children]' + ' ...' + convert(varchar, getdate(), 121)
GO
CREATE VIEW [dbo].[vDependency_Children] AS
/*
--------------------------------------------------------------
(c) 2019 David Berglin 
This file is part of the PowerBiVisibility project.
PowerBiVisibility is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
PowerBiVisibility is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with PowerBiVisibility.  If not, see https://www.gnu.org/licenses/.
--------------------------------------------------------------
Purpose is to Find Dependencies (Children) by Visuals/Measures (who are Parents)
        This view will show SSAS Measure and Column dependencies 

 populate [dbo].[Dependencies] using this Powershell scripts: 
    PowerBI Visibility.ps1        ... use SQL output option

 research...
    select * from [dbo].[Dependencies] 
    select * from [dbo].[Dependencies] where ChildAddress like '%[[]Total Hours]%'  -- [[]text] handles square bracket in "Like" clause

 Sample Script
    select distinct *
        ,'Select * from [dbo].[vDependency_Children] where TopMostParent   = '''+Replace(ParentAddress, '''', '''''')+'''' as [Script To Find Children Dependencies]
        ,'Select * from [dbo].[vDependency_Parents]  where DependencyChild = '''+Replace(ParentAddress, '''', '''''')+'''' as [Script To Find Parents Who Depend on Me]
    from [dbo].[vDependency_Children]
    order by ParentAddress

 Distinct Script
    select distinct    
         TopMostParent
        ,ChildAddress
    from [dbo].[vDependency_Children]

--------------------------------------------------------------
*/


with cte_nested as ( -- holds nested dependency objects

    -- Select the 'normal' (not recursive, just top level) rows of data
    select 
         1 as Recursion
        ,d.ParentAddress as TopMostParent
        ,Replace(d.ParentAddress, '''', '''''') as ScriptableParent -- replace single quotes with two single quotes... for pasting into SQL Select Scripts
        ,Replace(d.ChildAddress, '''', '''''')  as ScriptableChild -- replace single quotes with two single quotes... for pasting into SQL Select Scripts
        ,d.DependencyId
        ,d.ParentType
        ,d.ParentLocation
        ,d.ParentName
        ,d.ParentAddress
        ,d.ChildAddress
        ,d.ChildType 
        ,convert(varchar(5000), d.Source 
            + CHAR(13) + CHAR(10) + ' => ' + d.ParentLocation 
            + CHAR(13) + CHAR(10) + ' => ' + d.ParentName 
            + isnull(CHAR(13) + CHAR(10) + ' => ' + d.ChildAddress,'')) as [Trail]
        ,d.Content
    from [dbo].[Dependencies] d 
    where d.ChildAddress is not null -- no need to pull in null dependencies
        and d.ChildAddress <> ''

    union all 

    -- Recursively reach back into this same CTE to find nested data
    select 
         n.Recursion + 1
        ,n.TopMostParent
        ,Replace(d.ParentAddress, '''', '''''') as ScriptableParent
        ,Replace(d.ChildAddress, '''', '''''')  as ScriptableChild
        ,d.DependencyId
        ,d.ParentType
        ,d.ParentLocation
        ,d.ParentName
        ,d.ParentAddress
        ,d.ChildAddress
        ,d.ChildType
        ,convert(varchar(5000), n.[Trail] + isnull(CHAR(13) + CHAR(10) + ' => ' + d.ChildAddress, ''))
        ,d.Content
    from [dbo].[Dependencies] d
    join cte_nested n
        on  n.ChildAddress      =   d.ParentAddress -- pull in the child
        and d.ParentAddress     <> ''
        and n.ChildAddress      <> ''
        and n.Recursion         <   20
)

----------------------------------------------------------------
-- Get final results
----------------------------------------------------------------
    select distinct
         c.TopMostParent
        ,c.ParentType
        ,c.ParentLocation
        ,c.ParentName
        ,c.ParentAddress
        ,c.ChildType
        ,c.ChildAddress
        ,c.Recursion
        ,c.[Trail] + CHAR(13) + CHAR(10) + isnull(c.Content, '') + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) as [Trail]
        , CASE WHEN ChildAddress = '' then ''
               else Replace(Replace(s.Scripts, '[ReplaceWithName]', ChildAddress),  '[ReplaceWithScriptableName]', ScriptableChild) end as [Child Scripts]
        , Replace(Replace(s.Scripts,           '[ReplaceWithName]', ParentAddress), '[ReplaceWithScriptableName]', ScriptableParent)    as [Parent Scripts]
    from cte_nested c 
    join vDependency_Scripts s on 1 = 1 -- no real join needed, this copies the simple text of the script generator, and replaces values to produce scripts you can copy/paste


GO
    print 'finished view [dbo].[vDependency_Children]' + ' ...' + convert(varchar, getdate(), 121)




GO
IF Not Exists(
            -- Check for non PK Indexes
            SELECT * 
            FROM sys.indexes 
            WHERE object_id = OBJECT_ID('dbo.Dependencies') 
              and type_desc <> 'CLUSTERED'
)
begin
    print 'creating indexes on Dependency Tables' + ' ...' + convert(varchar, getdate(), 121)


    CREATE NONCLUSTERED INDEX [idx_Dependency_ParentChildIdPType] ON [dbo].[Dependencies]
    (
	    [ParentAddress] ASC,
	    [ChildAddress] ASC,
	    [DependencyId] ASC,
	    [ParentType] ASC
    )
    INCLUDE ( 	[Source],
	    [ParentLocation],
	    [ParentName],
	    [ChildType]) WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF) ON [PRIMARY]


    CREATE NONCLUSTERED INDEX [idx_Dependency_ChildParentIdPType] ON [dbo].[Dependencies]
    (
	    [ChildAddress] ASC,
	    [ParentAddress] ASC,
	    [DependencyId] ASC,
	    [ParentType] ASC
    )
    INCLUDE ( 	[Source],
	    [ParentLocation],
	    [ParentName],
	    [ChildType]) WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF) ON [PRIMARY]


    CREATE NONCLUSTERED INDEX [idx_Dependency_ChildParentCType] ON [dbo].[Dependencies]
    (
	    [ChildAddress] ASC,
	    [ParentAddress] ASC,
	    [ChildType] ASC
    )
    INCLUDE ( 	[ParentType],
	    [ChildLocation],
	    [ChildName]) WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF) ON [PRIMARY]

    print 'finished indexes on Dependency Tables' + ' ...' + convert(varchar, getdate(), 121)
end
else
begin
    print 'exists:  indexes on Dependency Tables' + ' ...' + convert(varchar, getdate(), 121)
end
GO


