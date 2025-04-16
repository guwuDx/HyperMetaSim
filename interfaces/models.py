from pydantic import BaseModel, Field
from typing import Dict, Optional


class GenericParams(BaseModel):
    height: float
    period: float
    e_phi: float # = Field(alias='e_phi')
    e_theta: float # = Field(alias='e_theta')

class ParamPayload(BaseModel):
    generic_parameters: GenericParams
    shape_parameters: Dict[str, float]

class Material(BaseModel):
    substrate_material: str # = Field(alias='substrate_material')
    pillar_material: str # = Field(alias='pillar_material')

class ExcitingSource(BaseModel):
    circular_polarization: Optional[bool] = False
    port: str
    mode: str

class DBOverwrite(BaseModel):
    db_host: str
    db_port: int
    db_user: str
    db_password: str
    db_name: str


class SingleSimRequest(BaseModel):
    shape_type: str
    wavelength_min: float
    wavelength_max: float
    material: Material
    exciting_source: ExcitingSource
    params: ParamPayload
    db_overwrite: Optional[DBOverwrite] = None
    redirect_to_db: bool