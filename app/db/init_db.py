from app.db.database import engine, Base
from app.model import entity

def init_db():
    Base.metadata.create_all(bind=engine)

if __name__ == "__main__":
    init_db()
