import sqlite3
conn = sqlite3.connect('database.db')
c = conn.cursor()

# Удаляем дубли категорий (оставляем одну запись для каждой пары name+season)
c.execute('DELETE FROM categories WHERE id NOT IN (SELECT MIN(id) FROM categories GROUP BY name, season)')
# Удаляем дубли туров
c.execute('DELETE FROM tours WHERE id NOT IN (SELECT MIN(id) FROM tours GROUP BY number, season)')
# Удаляем дубли клубов
c.execute('DELETE FROM clubs WHERE id NOT IN (SELECT MIN(id) FROM clubs GROUP BY name, season)')

conn.commit()

print('Категорий:', c.execute('SELECT COUNT(*) FROM categories').fetchone()[0])
print('Туров:', c.execute('SELECT COUNT(*) FROM tours').fetchone()[0])
print('Клубов:', c.execute('SELECT COUNT(*) FROM clubs').fetchone()[0])

conn.close()