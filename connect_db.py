import psycopg2
from config import load_config


# Подключаемся к PostgreSQL
connection = psycopg2.connect(
    host=load_config().get('host'),
    user=load_config().get('user'),
    password=load_config().get('password'),
    database=load_config().get('db_name')
)
connection.autocommit = True
# with connection.cursor() as cursor:
#     cursor.execute(
#         """
#         CREATE TABLE users(
#         id bigserial PRIMARY KEY,
#         user_id integer NOT NULL,
#         GX varchar(8) DEFAULT NULL,
#         status boolean DEFAULT NULL,
#         lan varchar(15) NOT NULL
#         );
#         """
#     )



def get_subscriptions(status=True):
    with connection.cursor() as cursor:
        cursor.execute(
            f"SELECT * FROM users WHERE status={status};"
        )
        return cursor.fetchall()


def get_by_gx(gx):
    with connection.cursor() as cursor:
        cursor.execute(
            f"SELECT * FROM users WHERE gx={gx};"
        )
        return cursor.fetchall()


def subscriber_exists(user_id, gx):
    """Проверяем существует ли пользователь"""
    with connection.cursor() as cursor:
        cursor.execute(
            f"SELECT * FROM users WHERE user_id={user_id} AND gx='{gx}';"
        )
        try:
            cursor.fetchall()[0]
            return True
        except IndexError:
            return False


def subscriber_status(user_id):
    """Проверяем активен ли пользователь"""
    with connection.cursor() as cursor:
        cursor.execute(
            f"SELECT status FROM users WHERE user_id={user_id};"
        )
        return cursor.fetchone()[0]


def add_subscriber(user_id, gx, status=True):
    """Добавляем пользователя"""
    with connection.cursor() as cursor:
        cursor.execute(
            f"INSERT INTO users (user_id, gx, status) VALUES ({user_id}, '{gx}', {status})"
        )   


def update_subscription(user_id, status):
    """Обновляем статус подписки"""
    with connection.cursor() as cursor:
        cursor.execute(
            f"UPDATE users SET status = {status} WHERE user_id = {user_id}"
        )

def get_len(user_id):
    with connection.cursor() as cursor:
        cursor.execute(
            f"SELECT lan FROM users WHERE user_id = {user_id};"
        )
        return cursor.fetchall()

def add_len(user_id, lan):
    with connection.cursor() as cursor:
            cursor.execute(
                f"INSERT INTO users (user_id, lan) VALUES ({user_id}, '{lan}')"
            )   

