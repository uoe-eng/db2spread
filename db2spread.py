from dataclasses import dataclass
import sqlalchemy as sa

@dataclass
class DB2Spread:
    engine: sa.Engine.Engine