-- ==================================================================================
-- 数据库迁移脚本: 为频率响应表添加 Mag 和 Phase 字段
-- 创建日期: 2025-08-20
-- 描述: 为所有 freq_resp 表添加振幅(Mag)和相位(Phase)字段，并为现有数据计算这些值
-- ==================================================================================

-- 设置 SQL 模式，确保安全操作
SET SQL_MODE = 'STRICT_TRANS_TABLES,ERROR_FOR_DIVISION_BY_ZERO,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION';

-- 开始事务
START TRANSACTION;

-- ==================================================================================
-- 1. 为 NIR (近红外) 频率响应表添加字段
-- ==================================================================================

-- CuboidPillar_freq_resp_NIR
ALTER TABLE CuboidPillar_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquarePillar_freq_resp_NIR
ALTER TABLE SquarePillar_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cylinder_freq_resp_NIR
ALTER TABLE Cylinder_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cross_freq_resp_NIR
ALTER TABLE Cross_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareHole_freq_resp_NIR
ALTER TABLE SquareHole_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SymmetricCross_freq_resp_NIR
ALTER TABLE SymmetricCross_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareRing_freq_resp_NIR
ALTER TABLE SquareRing_freq_resp_NIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- ==================================================================================
-- 2. 为 MIR (中红外) 频率响应表添加字段
-- ==================================================================================

-- CuboidPillar_freq_resp_MIR
ALTER TABLE CuboidPillar_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquarePillar_freq_resp_MIR
ALTER TABLE SquarePillar_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cylinder_freq_resp_MIR
ALTER TABLE Cylinder_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cross_freq_resp_MIR
ALTER TABLE Cross_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareHole_freq_resp_MIR
ALTER TABLE SquareHole_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SymmetricCross_freq_resp_MIR
ALTER TABLE SymmetricCross_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareRing_freq_resp_MIR
ALTER TABLE SquareRing_freq_resp_MIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- ==================================================================================
-- 3. 为 FIR (远红外) 频率响应表添加字段
-- ==================================================================================

-- CuboidPillar_freq_resp_FIR
ALTER TABLE CuboidPillar_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquarePillar_freq_resp_FIR
ALTER TABLE SquarePillar_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cylinder_freq_resp_FIR
ALTER TABLE Cylinder_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- Cross_freq_resp_FIR
ALTER TABLE Cross_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareHole_freq_resp_FIR
ALTER TABLE SquareHole_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SymmetricCross_freq_resp_FIR
ALTER TABLE SymmetricCross_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- SquareRing_freq_resp_FIR
ALTER TABLE SquareRing_freq_resp_FIR 
ADD COLUMN IF NOT EXISTS `Mag` DOUBLE NULL COMMENT '幅度' AFTER `imag_pt`,
ADD COLUMN IF NOT EXISTS `Phase` DOUBLE NULL COMMENT '相位' AFTER `Mag`;

-- ==================================================================================
-- 4. 创建触发器，自动计算 Mag 和 Phase 字段
-- ==================================================================================

-- 删除已存在的触发器（如果有的话）
DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_NIR_update;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_NIR_insert;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_NIR_update;

DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_MIR_update;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_MIR_insert;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_MIR_update;

DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_CuboidPillar_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_SquarePillar_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_Cylinder_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_Cross_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_SquareHole_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_SymmetricCross_freq_resp_FIR_update;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_FIR_insert;
DROP TRIGGER IF EXISTS tr_SquareRing_freq_resp_FIR_update;

-- NIR 表触发器
DELIMITER $$

-- CuboidPillar_freq_resp_NIR
CREATE TRIGGER tr_CuboidPillar_freq_resp_NIR_insert 
BEFORE INSERT ON CuboidPillar_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_CuboidPillar_freq_resp_NIR_update 
BEFORE UPDATE ON CuboidPillar_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquarePillar_freq_resp_NIR
CREATE TRIGGER tr_SquarePillar_freq_resp_NIR_insert 
BEFORE INSERT ON SquarePillar_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquarePillar_freq_resp_NIR_update 
BEFORE UPDATE ON SquarePillar_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cylinder_freq_resp_NIR
CREATE TRIGGER tr_Cylinder_freq_resp_NIR_insert 
BEFORE INSERT ON Cylinder_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cylinder_freq_resp_NIR_update 
BEFORE UPDATE ON Cylinder_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cross_freq_resp_NIR
CREATE TRIGGER tr_Cross_freq_resp_NIR_insert 
BEFORE INSERT ON Cross_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cross_freq_resp_NIR_update 
BEFORE UPDATE ON Cross_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareHole_freq_resp_NIR
CREATE TRIGGER tr_SquareHole_freq_resp_NIR_insert 
BEFORE INSERT ON SquareHole_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareHole_freq_resp_NIR_update 
BEFORE UPDATE ON SquareHole_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SymmetricCross_freq_resp_NIR
CREATE TRIGGER tr_SymmetricCross_freq_resp_NIR_insert 
BEFORE INSERT ON SymmetricCross_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SymmetricCross_freq_resp_NIR_update 
BEFORE UPDATE ON SymmetricCross_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareRing_freq_resp_NIR
CREATE TRIGGER tr_SquareRing_freq_resp_NIR_insert 
BEFORE INSERT ON SquareRing_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareRing_freq_resp_NIR_update 
BEFORE UPDATE ON SquareRing_freq_resp_NIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- MIR 表触发器
-- CuboidPillar_freq_resp_MIR
CREATE TRIGGER tr_CuboidPillar_freq_resp_MIR_insert 
BEFORE INSERT ON CuboidPillar_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_CuboidPillar_freq_resp_MIR_update 
BEFORE UPDATE ON CuboidPillar_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquarePillar_freq_resp_MIR
CREATE TRIGGER tr_SquarePillar_freq_resp_MIR_insert 
BEFORE INSERT ON SquarePillar_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquarePillar_freq_resp_MIR_update 
BEFORE UPDATE ON SquarePillar_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cylinder_freq_resp_MIR
CREATE TRIGGER tr_Cylinder_freq_resp_MIR_insert 
BEFORE INSERT ON Cylinder_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cylinder_freq_resp_MIR_update 
BEFORE UPDATE ON Cylinder_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cross_freq_resp_MIR
CREATE TRIGGER tr_Cross_freq_resp_MIR_insert 
BEFORE INSERT ON Cross_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cross_freq_resp_MIR_update 
BEFORE UPDATE ON Cross_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareHole_freq_resp_MIR
CREATE TRIGGER tr_SquareHole_freq_resp_MIR_insert 
BEFORE INSERT ON SquareHole_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareHole_freq_resp_MIR_update 
BEFORE UPDATE ON SquareHole_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SymmetricCross_freq_resp_MIR
CREATE TRIGGER tr_SymmetricCross_freq_resp_MIR_insert 
BEFORE INSERT ON SymmetricCross_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SymmetricCross_freq_resp_MIR_update 
BEFORE UPDATE ON SymmetricCross_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareRing_freq_resp_MIR
CREATE TRIGGER tr_SquareRing_freq_resp_MIR_insert 
BEFORE INSERT ON SquareRing_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareRing_freq_resp_MIR_update 
BEFORE UPDATE ON SquareRing_freq_resp_MIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- FIR 表触发器
-- CuboidPillar_freq_resp_FIR
CREATE TRIGGER tr_CuboidPillar_freq_resp_FIR_insert 
BEFORE INSERT ON CuboidPillar_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_CuboidPillar_freq_resp_FIR_update 
BEFORE UPDATE ON CuboidPillar_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquarePillar_freq_resp_FIR
CREATE TRIGGER tr_SquarePillar_freq_resp_FIR_insert 
BEFORE INSERT ON SquarePillar_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquarePillar_freq_resp_FIR_update 
BEFORE UPDATE ON SquarePillar_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cylinder_freq_resp_FIR
CREATE TRIGGER tr_Cylinder_freq_resp_FIR_insert 
BEFORE INSERT ON Cylinder_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cylinder_freq_resp_FIR_update 
BEFORE UPDATE ON Cylinder_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- Cross_freq_resp_FIR
CREATE TRIGGER tr_Cross_freq_resp_FIR_insert 
BEFORE INSERT ON Cross_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_Cross_freq_resp_FIR_update 
BEFORE UPDATE ON Cross_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareHole_freq_resp_FIR
CREATE TRIGGER tr_SquareHole_freq_resp_FIR_insert 
BEFORE INSERT ON SquareHole_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareHole_freq_resp_FIR_update 
BEFORE UPDATE ON SquareHole_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SymmetricCross_freq_resp_FIR
CREATE TRIGGER tr_SymmetricCross_freq_resp_FIR_insert 
BEFORE INSERT ON SymmetricCross_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SymmetricCross_freq_resp_FIR_update 
BEFORE UPDATE ON SymmetricCross_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

-- SquareRing_freq_resp_FIR
CREATE TRIGGER tr_SquareRing_freq_resp_FIR_insert 
BEFORE INSERT ON SquareRing_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

CREATE TRIGGER tr_SquareRing_freq_resp_FIR_update 
BEFORE UPDATE ON SquareRing_freq_resp_FIR
FOR EACH ROW
BEGIN
    IF NEW.real_pt IS NOT NULL AND NEW.imag_pt IS NOT NULL THEN
        SET NEW.Mag = SQRT(POW(NEW.real_pt, 2) + POW(NEW.imag_pt, 2));
        SET NEW.Phase = ATAN2(NEW.imag_pt, NEW.real_pt);
    END IF;
END$$

DELIMITER ;

-- ==================================================================================
-- 5. 更新现有数据，计算 Mag 和 Phase 字段
-- ==================================================================================

-- NIR 表数据更新
UPDATE CuboidPillar_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquarePillar_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cylinder_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cross_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareHole_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SymmetricCross_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareRing_freq_resp_NIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

-- MIR 表数据更新
UPDATE CuboidPillar_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquarePillar_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cylinder_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cross_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareHole_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SymmetricCross_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareRing_freq_resp_MIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

-- FIR 表数据更新
UPDATE CuboidPillar_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquarePillar_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cylinder_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE Cross_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareHole_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SymmetricCross_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

UPDATE SquareRing_freq_resp_FIR 
SET Mag = SQRT(POW(real_pt, 2) + POW(imag_pt, 2)),
    Phase = ATAN2(imag_pt, real_pt)
WHERE real_pt IS NOT NULL AND imag_pt IS NOT NULL;

-- ==================================================================================
-- 6. 验证数据迁移结果
-- ==================================================================================

-- 检查各表的字段是否已添加
SELECT 
    'CuboidPillar_freq_resp_NIR' as table_name,
    COUNT(*) as total_records,
    COUNT(Mag) as mag_not_null,
    COUNT(Phase) as phase_not_null,
    COUNT(CASE WHEN real_pt IS NOT NULL AND imag_pt IS NOT NULL AND (Mag IS NULL OR Phase IS NULL) THEN 1 END) as missing_calculations
FROM CuboidPillar_freq_resp_NIR

UNION ALL

SELECT 
    'SquarePillar_freq_resp_FIR' as table_name,
    COUNT(*) as total_records,
    COUNT(Mag) as mag_not_null,
    COUNT(Phase) as phase_not_null,
    COUNT(CASE WHEN real_pt IS NOT NULL AND imag_pt IS NOT NULL AND (Mag IS NULL OR Phase IS NULL) THEN 1 END) as missing_calculations
FROM SquarePillar_freq_resp_FIR;

-- 提交事务
COMMIT;

-- ==================================================================================
-- 迁移完成
-- ==================================================================================
SELECT 'Migration completed successfully!' as status;
