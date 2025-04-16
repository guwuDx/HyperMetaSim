CREATE INDEX idx_gp_query ON generic_parameters (
  s_param_id,
  substrate_material_id,
  pillar_material_id,
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