import yagmail
pdf = ('D:\\Dev_autonomous\\Python\\Email_sender\\report.pdf')
yag = yagmail.SMTP('belukrprom.bot@gmail.com','belprom1')
contents = [
    "test yagmail",
    "You can find an audio file attached.", pdf
]

to_email = ('test')
yag.send(to_email, 'test', contents)
