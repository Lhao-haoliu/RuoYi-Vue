-- Portal Phase 1: menu + permission bootstrap for RuoYi-Vue3 frontend.
-- Run after base schema/data script (e.g. ry_20250522.sql).

-- 1) Hide built-in demo/tool modules.
UPDATE sys_menu SET visible = '1', status = '1' WHERE menu_id IN (2, 3, 4);
UPDATE sys_menu SET visible = '1', status = '1' WHERE parent_id IN (2, 3, 4);
UPDATE sys_menu
SET visible = '1', status = '1'
WHERE parent_id IN (
    SELECT menu_id
    FROM (SELECT menu_id FROM sys_menu WHERE parent_id IN (2, 3, 4)) t
);

-- 2) Keep role-menu clean: remove disabled menu assignments from role `common` (role_id=2).
DELETE rm
FROM sys_role_menu rm
JOIN sys_menu m ON m.menu_id = rm.menu_id
WHERE rm.role_id = 2
  AND m.status = '1';

-- 3) Portal top-level menus.
INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2000, '报表中心', 0, 2, 'report', NULL, '', '', 1, 0, 'M', '0', '0', '', 'chart', 'admin', NOW(), '', NULL, 'Portal报表目录'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2000);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2003, 'Commander', 0, 3, 'commander', NULL, '', '', 1, 0, 'M', '0', '0', '', 'guide', 'admin', NOW(), '', NULL, 'Portal指挥目录'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2003);

-- 4) Portal C-type menus (menu visibility controlled by role + perms).
INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2001, '任务明细', 2000, 1, 'task-list', 'report/task-list', '', '', 1, 0, 'C', '0', '0', 'report:view', 'list', 'admin', NOW(), '', NULL, '任务明细菜单'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2001);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2002, '统计报表', 2000, 2, 'statistics', 'report/statistics', '', '', 1, 0, 'C', '0', '0', 'report:export', 'form', 'admin', NOW(), '', NULL, '统计报表菜单'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2002);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2004, '任务启动', 2003, 1, 'start', 'commander/start', '', '', 1, 0, 'C', '0', '0', 'commander:start', 'job', 'admin', NOW(), '', NULL, '任务启动菜单'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2004);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2005, '公版图管理', 2003, 2, 'public-map', 'commander/public-map', '', '', 1, 0, 'C', '0', '0', 'commander:map', 'tree', 'admin', NOW(), '', NULL, '公版图管理菜单'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2005);

-- 5) Permission points as F-type menu permissions.
INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2010, '报表查看', 2001, 1, '', '', '', '', 1, 0, 'F', '0', '0', 'report:view', '#', 'admin', NOW(), '', NULL, '报表查看权限点'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2010);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2011, '报表导出', 2002, 1, '', '', '', '', 1, 0, 'F', '0', '0', 'report:export', '#', 'admin', NOW(), '', NULL, '报表导出权限点'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2011);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2012, '任务启动', 2004, 1, '', '', '', '', 1, 0, 'F', '0', '0', 'commander:start', '#', 'admin', NOW(), '', NULL, 'Commander启动权限点'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2012);

INSERT INTO sys_menu
    (menu_id, menu_name, parent_id, order_num, path, component, `query`, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_by, create_time, update_by, update_time, remark)
SELECT 2013, '公版图管理', 2005, 1, '', '', '', '', 1, 0, 'F', '0', '0', 'commander:map', '#', 'admin', NOW(), '', NULL, 'Commander地图权限点'
WHERE NOT EXISTS (SELECT 1 FROM sys_menu WHERE menu_id = 2013);
