CREATE INDEX idx_gp_query ON generic_parameters (
  substrate_material_id,
  pillar_material_id,
  s_param_id,
  period,
  height,
  thickness,
  e_theta
);

CREATE INDEX idx_freq_gp_shp ON `Cross_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `Cross_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `Cross_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `CuboidPillar_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `CuboidPillar_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `CuboidPillar_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `Cylinder_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `Cylinder_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `Cylinder_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `SquareHole_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquareHole_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquareHole_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `SquarePillar_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquarePillar_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquarePillar_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `SquareRing_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquareRing_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SquareRing_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);

CREATE INDEX idx_freq_gp_shp ON `SymmetricCross_freq_resp_FIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SymmetricCross_freq_resp_MIR` (Frequency, GP_ID, ShP_ID);
CREATE INDEX idx_freq_gp_shp ON `SymmetricCross_freq_resp_NIR` (Frequency, GP_ID, ShP_ID);



ALTER TABLE `CuboidPillar_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `CuboidPillar_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `CuboidPillar_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `SquarePillar_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquarePillar_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquarePillar_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `Cylinder_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `Cylinder_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `Cylinder_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `Cross_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `Cross_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `Cross_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `SquareHole_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquareHole_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquareHole_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `SymmetricCross_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SymmetricCross_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SymmetricCross_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);

ALTER TABLE `SquareRing_freq_resp_FIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquareRing_freq_resp_MIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
ALTER TABLE `SquareRing_freq_resp_NIR` ADD INDEX idx_gp_freq (GP_ID, Frequency);
