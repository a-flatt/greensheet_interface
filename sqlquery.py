import psycopg2

def fetchR2(code_range, job_id):

	print('testing... {}'.format(job_id))
	conn = psycopg2.connect("dbname=buildsoftStandalone user=postgres port=5432")
	conn.set_session(readonly = True)
	cur = conn.cursor()
	cur.execute("""
                SELECT
					codetext AS costcode,	
					codes.codedescription,
					SUM (markeduptotal) AS total
				FROM 
					tradeitemrates
				INNER JOIN tradenodes ON tradenodes.id = tradeitemrates.ownerid
				INNER JOIN traderatessortcodes ON tradeitemrates.id = traderatessortcodes.rate_id
				INNER JOIN codes ON traderatessortcodes.codetext = codes.code
				INNER JOIN groupcodes ON codes.groupid = groupcodes.groupid
				INNER JOIN jobs ON tradenodes.jobid = jobs.job_id
				WHERE jobs.code = '{}' AND groupcodes.groupcode = 'RCC'
				GROUP BY costcode, codes.codedescription;
                """.format(job_id.upper()))
	costlist = list(cur.fetchall())
	cur.close()
	conn.close()
	return [row for row in costlist if int(row[0]) > code_range[0] and int(row[0]) < code_range[1]]

 
