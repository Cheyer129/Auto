import sqlite3

conn = sqlite3.connect('payoffs.db')

# create cursor
c = conn.cursor()

# # create a table
# c.execute("""CREATE TABLE payoffs (
#         borrower text,
#         loan_number text,
#         heloc_number text,
#         loan_type text,
#         contact text,
#         contact_email text,
#         phone_number text,
#         residents text,
#         street_address text,
#         city_state_zip text
#     )""")

c.execute("INSERT INTO payoffs VALUES ('Heyer', '1122334455', '', 'CEMA', 'Chris Heyer,Esq.', 'chrisheyer0@gmail.com', '5166609807', 'Chris Heyer', '45-35 46th St, 5E', 'Woodside, NY 11377')")
# c.execute("DELETE FROM payoffs WHERE rowid > 1")

conn.commit()

# QUERY THE DATABASE
c.execute("SELECT rowid, * FROM payoffs")
# print(c.fetchone())
# print(c.fetchmany(3))
# c.fetchall()
print(c.fetchall())


# searching the database for specific shit
# c.execute("SELECT rowid, * FROM customers WHERE last_name = 'Stupid'")


# update our records
# c.execute("""UPDATE customers SET first_name =

# items = c.fetchall()
# # print(items)
# for item in items:
#     print(item)

# commit our command
# conn.commit()

# close our connection
conn.close()
