import os
import glob
import datetime
import mailbox
from email.header import decode_header

# 過去取得フラグ
get_all = False
#get_all = True

# Thunderbirdのパス
root = os.environ['USERPROFILE'] + "\AppData\Roaming\Thunderbird\Profiles"
files = glob.glob(root + "\**", recursive=True)
#print(root)

# デスクトップパス
if os.name == 'nt':
	home = os.getenv('USERPROFILE')
else:
	home = os.getenv('HOME')
desktop_dir = os.path.join(home, 'Desktop')

# 書き込みファイルオープン
if get_all:
	write_file_name = "メール情報-ALL.sql"
else:
	write_file_name = "メール情報.sql"
wfn = os.path.join(desktop_dir, write_file_name)
fw = open(wfn, "w", errors="ignore")

# 前回実行日時
last_time_file = "last_time.txt"
is_last_time_file = os.path.isfile(last_time_file)
now_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
if is_last_time_file:
	flt = open(last_time_file, 'r')
	dt_last_time = flt.read()
	flt.close()
else:
	dt_last_time = now_time

# 実行時間書き込み
flt = open(last_time_file, "w", errors="ignore")
flt.write(now_time)
flt.close()

print('前回実行日時：' + dt_last_time)

# カウンター
prog = 0

# 読み込みファイル
#dir = "C:\\Users\\jinpachi\\AppData\\Roaming\\Thunderbird\\Profiles\\0i41g78k.default\\Mail\\Local Folders\\Inbox"
#mail_box = mailbox.mbox(dir)

# ヘッダを取得
def get_header(msg, name):
	header = ''
	if msg[name]:
		for tup in decode_header(str(msg[name])):
			if type(tup[0]) is bytes:
				charset = tup[1]
				if charset:
					try:
						header += tup[0].decode(tup[1])
					except:
#						print(str(msg[name]))
						header = '変換エラー'
				else:
					header += tup[0].decode()
			elif type(tup[0]) is str:
				header += tup[0]
	return header

# 本文を取得
def get_content(msg):
	if msg.is_multipart() is False:
		# SinglePart

		# ファイル名の取得
#		attach_fname = msg.get_filename()
#		print(attach_fname)

		payload = msg.get_payload(decode=True)	# 備考の※1
		charset = msg.get_content_charset()		# 備考の※2
		if charset is not None:
			payload = payload.decode(charset, "ignore")
			return payload
	else:
		# MultiPart
		for part in msg.walk():
			payload = part.get_payload(decode=True)
			if payload is None:
				continue
			charset = part.get_content_charset()
			if charset == 'cp-850':
				payload = payload.decode("UTF-8", "ignore")
				return payload
			elif charset is not None:
				payload = payload.decode(charset, "ignore")
				return payload

# 日付変換
def get_date(str):
	res = ''

	try:
		#Wed, 2 Jul 2022 03:11:06 +0900
		res = datetime.datetime.strptime(str, '%a, %d %b %Y %H:%M:%S %z').strftime('%Y/%m/%d %H:%M:%S')
	except:
		try:
			#2 Jul 2022 03:11:06 +0900
			res = datetime.datetime.strptime(str, '%d %b %Y %H:%M:%S %z').strftime('%Y/%m/%d %H:%M:%S')
		except:
			try:
				#:15 Nov 2019 03:16:04 +0900
				res = datetime.datetime.strptime(str, ':%d %b %Y %H:%M:%S %z').strftime('%Y/%m/%d %H:%M:%S')
			except:
				pass

	if res == '':
		print('Date Error:'+ str)
		res = datetime.datetime.now().strftime('%Y-%m-%d 00:00:00')
		print(res)

	if res == '1899/12/30 00:00:00':
		res = datetime.datetime.now().strftime('%Y-%m-%d 00:00:00')
		print(res)

	return res

for fp in files :

	# フォルダーは無視
	if os.path.isdir(fp):
		continue

	# 拡張子を取得
	ext = root, ext = os.path.splitext(fp)

	# 拡張子なしのファイルのみ対象
	if ext != '':
		continue

	# ファイル名を取得（拡張子あり）
	fn = os.path.basename(fp)

	# 除外ファイル
	if fn.lower() == 'sent':
		continue
	elif fn.lower() == 'trash':
		continue
	elif fn.lower() == 'drafts':
		continue
	elif fn.lower() == 'junk':
		continue

	# Inboxファイル以外は除外
	if fn.lower() != 'inbox':
		continue

	# サブディレクトリーに分類されたメール
	if '.sbd' in fp:
		if 'sent.sbd' in fp.lower():
			# 送信済みフォルダーは無視
			continue
		elif 'trash.sbd' in fp.lower():
			# 送信済みフォルダーは無視
			continue
#		else:
#			print(fp)

	# ファイル読み込み
	print('\n「{0}」を読み込み中です。'.format(fp))
	mail_box = mailbox.mbox(fp)

	tbl1 = str.maketrans({'\"' : None})
	tbl2 = str.maketrans({'\'' : None})
	tbl3 = str.maketrans({'\n' : '\\n'})
	tbl4 = str.maketrans({'\r' : '\\n'})
	for key in mail_box.keys():
		# 進捗
		prog = prog + 1
		print("\r[{0}]".format(prog), end="")
		# メッセージ取得
		msg = mail_box.get(key)
		# 件名
		subject = get_header(msg, 'Subject')
		if 'SPAM' in subject:
			continue
		subject = subject.translate(tbl1)
		subject = subject.translate(tbl2)
		subject = subject.translate(tbl3)
		subject = subject.translate(tbl4)
		# 受信日時
		rcv_time = get_header(msg, 'Date')
		rcv_time = rcv_time.replace(' (JST)','').replace(' (UTC)','').replace(' (CST)','').replace(' (GMT)','').replace(' (EEST)','').replace(' (CDC)','').replace(' (CDT)','')
		rcv_time2 = get_date(rcv_time)
		# 本日以前のメールは除外
		if not get_all and rcv_time2 < dt_last_time:
#			print('cut_day:' + rcv_time2)
			continue
		# Message-ID
		mid = get_header(msg, 'Message-ID')
		mid = mid.translate(tbl1)
		mid = mid.translate(tbl2)
		mid = mid.translate(tbl3)
		mid = mid.translate(tbl4)
#		print('Message-ID:' + mid)
		# 差出人
		sender = get_header(msg, 'From')
		sender = sender.translate(tbl1)
		sender = sender.translate(tbl2)
		sender = sender.translate(tbl3)
		sender = sender.translate(tbl4)
		# 受信者
		receiver = get_header(msg, 'To')
		receiver = receiver.translate(tbl1)
		receiver = receiver.translate(tbl2)
		receiver = receiver.translate(tbl3)
		receiver = receiver.translate(tbl4)
		# 本文
		content = get_content(msg)
		content = str(content)
		content = content.translate(tbl1)
		content = content.translate(tbl2)
		content = content.translate(tbl3)
		content = content.translate(tbl4)

		# 書き込み
#		sql = "REPLACE INTO `mailert`.`t_mail` (`message_id`,`time`,`subject`,`sender`,`receiver`,`msg`) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}');"
		sql = "INSERT INTO `mailert`.`t_mail` (`message_id`,`time`,`subject`,`sender`,`receiver`,`msg`) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')"
		sql += " ON DUPLICATE KEY UPDATE subject = VALUES (subject), sender = VALUES (sender), receiver = VALUES (receiver), msg = VALUES (msg);"

		sql = sql.format(mid ,rcv_time2 , subject ,sender, receiver, content) + '\n'
		fw.write(sql)

print('\n{0}件抽出しました。メール＋の「設定－受信ログ取り込み」で「{1}」を取り込んでください。'.format(prog, write_file_name))
