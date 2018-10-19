import pyodbc
import datetime
# connect_str = (
#     r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
#     r'UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;'
#     r'FIL={MS Access};DriverId=25;DefaultDir=T:/CITDR/CITSQ/QA/Operational/Status Reports/QA reviews/Access_DB;'
#     r'DBQ=T:/CITDR/CITSQ/QA/Operational/Status Reports/QA reviews/Access_DB/QA Project Review_Database_v04 - Update test.accdb;')

# # Create connections with Access
# conn = pyodbc.connect(connect_str)


default_dir = 'C:/Users/ashah13/Documents/'
dbq = 'C:/Users/ashah13/OneDrive - WBG/Desktop/Python_Files/QA Project Review_Database_v04 - Update test.mdb'
global cursor,conn

def setup_conn(default_dir,dbq):
    global cursor,conn
    conn_str = 'DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};DriverId=25;DefaultDir=%s;DBQ=%s;' % (default_dir,dbq)
    connect_str2 = (conn_str)
    conn = pyodbc.connect(connect_str2)
    cursor = conn.cursor()
    print("DB Connected")

'''
    Query
'''
def query_projects():
    sql = "SELECT project_id,project_name FROM Project"
    query_result = query_sql(sql)

    # print(query_result)

    return query_result

def query_teams():
    sql = "SELECT DISTINCT (t.team_name) FROM Teams as t;"
    query_result = query_sql(sql)

    # print(query_result)

    return query_result

def query_projects_by_team(team_name):
    sql = "SELECT p.project_id, p.project_name,t.team_name FROM Teams as t \
    INNER JOIN Project as p ON t.project_id = p.project_id WHERE t.team_name = '%s';" % (team_name)
    query_result = query_sql(sql)
    return query_result

def query_project_modules(project_name):
    # print("project_name",project_name)
    sql = "SELECT project_module_name FROM Project_Module LEFT JOIN Project ON Project.project_id = Project_Module.project_id WHERE Project.project_name = '%s';" % (project_name)
    # print(sql)
    query_result = query_sql(sql)

    # print(query_result)

    return query_result

def query_project_teams(project_name):
    # print("project_name",project_name)
    sql = "SELECT team_name FROM Project_Module LEFT JOIN Project ON Project.project_id = Project_Module.project_id WHERE Project.project_name = '%s';" % project_name
    sql = "SELECT rs.status_name,f.Findings_id,p.doc_name_version,m.project_module_name,fc.finding_name, f.QA_review_observation,  a.artifact_name, s.severity_name  FROM (((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    INNER JOIN Artifact_Type as a ON p.artifact_type_id = a.artifact_type_id ) \
    INNER JOIN Finding_Category as fc ON f.finding_category_id = fc.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Resolution_Status rs ON rs.status_id = f.status_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND pr.project_name = '%s' AND m.project_module_name = '%s' AND (f.status_id = 2 OR f.status_id = 3);" % (proj_name,mod_name)

    query_result = query_sql(sql)

    # print(query_result)

    return query_result




def query_projects_location(name):
    #sql = "SELECT QA_feedback_location FROM Project WHERE project_id = %s;" % (proj_id)
    sql = "SELECT QA_feedback_location FROM Project WHERE project_name = '%s';" % (name)

    # print(sql)

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find this project location in database')
        return ''
    
    location = query_result[0][0]

    return location

def query_projects_id(project_name):
    #sql = "SELECT QA_feedback_location FROM Project WHERE project_id = %s;" % (proj_id)
    sql = "SELECT project_id FROM Project WHERE project_name = '%s';" % (project_name)
    # print(sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find this project ID in database')
        return ''

    location = query_result[0][0]
    return location

def query_project_details(project_name):
    #sql = "SELECT QA_feedback_location FROM Project WHERE project_id = %s;" % (proj_id)
    sql = "SELECT * FROM Project WHERE project_name = '%s';" % (project_name)
    # print(sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find this project ID in database')
        return ''

    details_arr = query_result[0]
    return details_arr

def query_project_module_by_name(project_module_name):
    sql = "SELECT project_module_id FROM Project_Module WHERE project_module_name = '%s';" % (project_module_name)
    # print(sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find this project MODULE in database')
        return ''

    location = query_result[0][0]
    return location

def query_observation_type(artifact):
    sql = "SELECT finding_name FROM Finding_Category WHERE artifact_type_id = (select artifact_type_id from Artifact_type where artifact_name = '%s') and \
            new_version = True;" % (artifact)

    # print(sql)

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find observation_type in database')
        return ''
    
    types = query_result
    types = [ob[0] for ob in types]
    # print(types)

    return types

def query_resolution_status():
    sql = "SELECT * FROM Resolution_Status;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find status in database')
        return ''
    
    return query_result

def query_status_top():
    sql = "SELECT TOP 1 * FROM Resolution_Status ORDER BY status_id DESC;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find status top in database')
        return ''
    
    return query_result

def query_severity():
    sql = "SELECT * FROM Severity;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find severity in database')
        return ''
    
    return query_result

def query_severity_top():
    sql = "SELECT TOP 1 * FROM Severity ORDER BY severity_id DESC;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find severity top in database')
        return ''
    
    return query_result

def query_artifact():
    sql = "SELECT * FROM Artifact_Type;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find artifact in database')
        return ''
    
    return query_result

def query_artifact_name():
    sql = "SELECT artifact_name FROM Artifact_Type;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find artifact name in database')
        return ''
    
    result = []
    for x in query_result:
        result.append(x[0])

    return result


def query_finding_name():
    sql = "SELECT fc.finding_name FROM Finding_Category as fc WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%';"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find finding name in database')
        return ''
    
    result = []
    for x in query_result:
        result.append(x[0])

    return result

def query_team_id(team_name):
    sql = "SELECT team_id FROM Teams WHERE team_name = '%s';" % (team_name)
    # print(sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find this query_team_id in database')
        return ''

    location = query_result[0][0]
    return location

def query_team_name():
    sql = "SELECT DISTINCT t.team_name FROM Teams as t;"
    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find team name in database')
        return ''
    
    result = []
    for x in query_result:
        result.append(x[0])

    return result

def query_artifact_top():
    sql = "SELECT TOP 1 * FROM Artifact_Type ORDER BY artifact_type_id DESC;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find artifact top in database')
        return ''
    
    return query_result

def query_artifact_id(artifact_name):
    sql = "SELECT * FROM Artifact_Type WHERE artifact_name = '%s';" % (artifact_name)
    query_result = query_sql(sql)
    
    if query_result == []:
        print('Cannot find artifact in database')
        return ''
    
    return query_result[0][0]


def query_finding_category_id(finding_type):
    finding_type = finding_type.replace('TBD\'s','TBD')
    sql = "SELECT * FROM Finding_Category WHERE finding_name = '%s';" % (finding_type)
    query_result = query_sql(sql)
    
    if query_result == []:
        print('Cannot find artifact in database')
        return ''
    
    return query_result[0][0]

def query_status_id(status):
    sql = "SELECT * FROM Resolution_Status WHERE status_name = '%s';" % (status)
    query_result = query_sql(sql)
    
    if query_result == []:
        print('Cannot find artifact in database')
        return ''
    
    return query_result[0][0]

def query_severity_id(severity):
    sql = "SELECT * FROM Severity WHERE severity_name = '%s';" % (severity)
    query_result = query_sql(sql)
    
    if query_result == []:
        print('Cannot find artifact in database')
        return ''
    
    return query_result[0][0]


def query_finding():
    sql = "SELECT * FROM Finding_Category as fc WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%';"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find finding in database')
        return ''
    
    return query_result

def query_finding_top():
    sql = "SELECT TOP 1 * FROM Finding_Category as fc WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND ORDER BY fc.finding_category_id DESC;"

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find finding top in database')
        return ''
    
    return 

def query_findings_by_project(proj_name,mod_name):
    sql = "SELECT f.QA_review_observation,fc.finding_name,a.artifact_name,s.severity_name  FROM ((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND pr.project_name = '%s' AND m.project_module_name = '%s' AND f.status_id = 1 ;" % (proj_name,mod_name)

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find finding in database')
        return ''
    
    return query_result  

def query_findings_by_projects(team_name,proj_name,mod_name,artifact_name,status_id):
    status_list = ', '.join('%d' % (i) for i in status_id)
    team_list = ', '.join('\'%s\'' % (i) for i in team_name)
    artifact_list= ', '.join('\'%s\'' % (i) for i in artifact_name)
    sql = "SELECT f.QA_review_observation,fc.finding_name,a.artifact_name,s.severity_name FROM ((((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Resolution_Status rs ON rs.status_id = f.status_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id) \
    LEFT JOIN Teams as t ON pr.project_id = t.project_id  WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND (t.team_name IN (%s) OR t.team_name IS NULL) AND pr.project_name = '%s' AND m.project_module_name = '%s' \
    AND a.artifact_name IN (%s) AND f.status_id IN (%s);" % (team_list,proj_name,mod_name,artifact_list,status_list)

    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find findings by project and module in database')
        return ''
    
    return query_result  

def query_artifacts_details(team_name,proj_name,mod_name,status_id):
    status_list = ', '.join('%d' % (i) for i in status_id)
    team_list= ', '.join('\'%s\'' % (i) for i in team_name)
    sql = "SELECT DISTINCT a.artifact_name  FROM ((((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Resolution_Status rs ON rs.status_id = f.status_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id ) \
    LEFT JOIN Teams as t ON pr.project_id = t.project_id  WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND (t.team_name IN (%s) OR t.team_name IS NULL) AND pr.project_name = '%s' AND m.project_module_name = '%s' AND f.status_id IN (%s);" % (team_list,proj_name,mod_name,status_list)
    # print("artifact",sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find artifact details in database')
        return ''

    return query_result 

def query_findings(team_name,proj_name,mod_name,artifact_name,status_id):
    status_list = ', '.join('%d' % (i) for i in status_id)
    team_list = ', '.join('\'%s\'' % (i) for i in team_name)
    artifact_list = ', '.join('\'%s\'' % (i) for i in artifact_name)
    sql = "SELECT DISTINCT fc.finding_name  FROM ((((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Resolution_Status rs ON rs.status_id = f.status_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id ) \
    LEFT JOIN Teams as t ON pr.project_id = t.project_id  WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND (t.team_name IN (%s) OR t.team_name IS NULL) AND pr.project_name = '%s' AND m.project_module_name = '%s' \
    AND a.artifact_name IN (%s) AND f.status_id IN (%s);" % (team_list,proj_name,mod_name,artifact_list,status_list)
    # print("findings",sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find findings in database')
        return ''

    return query_result  
 

def query_findings_details(team_name,proj_name,mod_name,artifact_name,finding_name,status_id):
    status_list = ', '.join('%d' % (i) for i in status_id)
    artifact_list = ', '.join('\'%s\'' % (i) for i in artifact_name)
    finding_list = ', '.join('\'%s\'' % (i) for i in finding_name)
    team_list = ', '.join('\'%s\'' % (i) for i in team_name)
    sql = "SELECT DISTINCT (f.Findings_id), rs.status_name, fc.finding_name, p.doc_name_version, m.project_module_name, f.QA_review_observation, a.artifact_name, s.severity_name  FROM ((((((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Severity s ON f.severity_id = s.severity_id ) \
    INNER JOIN Resolution_Status rs ON rs.status_id = f.status_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id ) \
    LEFT JOIN Teams as t ON pr.project_id = t.project_id  \
    WHERE (t.team_name IN (%s) OR t.team_name IS NULL) AND pr.project_name = '%s' AND m.project_module_name = '%s' \
    AND a.artifact_name IN (%s) AND fc.finding_name IN (%s) AND f.status_id  IN (%s);" % (team_list,proj_name,mod_name,artifact_list,finding_list,status_list)
    # print("finding details",sql)
    query_result = query_sql(sql)
    if query_result == []:
        print('Cannot find finding details in database')
        return ''

    return query_result  


def query_findings_finding_id(finding_id):
    sql = "SELECT p.doc_name_version , p.count, pr.project_name, m.project_module_name, f.location_in_artifact, fc.finding_name, f.QA_review_observation, f.project_response, f.review_date, f.followup_review_date_1, f.followup_review_date_2, f.followup_review_date_3, f.followup_review_date_4 ,f.Findings_id, a.artifact_name FROM (((( Findings as f \
    INNER JOIN Project_Artifact as p ON f.PA_id = p.PA_id ) \
    LEFT JOIN Artifact_Type as a ON a.artifact_type_id = p.artifact_type_id ) \
    LEFT JOIN Finding_Category as fc ON fc.finding_category_id = f.finding_category_id) \
    INNER JOIN Project_Module m ON p.project_module_id = m.project_module_id ) \
    INNER JOIN Project pr ON pr.project_id = m.project_id WHERE fc.finding_name NOT LIKE '%%Clarification%%' AND fc.finding_name NOT LIKE '%%Tracking%%' AND f.Findings_id = %d ;" % (finding_id)

    query_result = query_sql(sql)

    if query_result == []:
        print('Cannot find finding in database for finding id %s' % str(finding_id))
        return ''
    
    return query_result  


'''
    Insert values 
    sql = "SELECT project_module_name FROM Project_Module LEFT JOIN Project ON Project.project_id = Project_Module.project_id WHERE Project.project_name = '%s';" % project_name

'''

def insert_project(project): # [[clarity_id,project_name,staff_project_mgr,QA_feedback_location]]
    # print(projects)
    count = 0
    print(project)
    sql = "INSERT INTO Project  (clarity_id, project_name, staff_project_mgr, QA_feedback_location, portfolio) \
            SELECT TOP 1 '%s', '%s', '%s', '%s',Null FROM Project\
            WHERE NOT EXISTS (SELECT * FROM Project WHERE project_name = '%s');" % (project[0],project[1],project[2],project[3],project[1])
    # print(sql)
    cursor.execute(sql)
    count += 1

    conn.commit()

    print(str(count) + ' project inserted')


def insert_proj_module(modules): # [[module_name,id]]
    
    count = 0
    for module in modules:
        print(module)
        count += 1
        sql = "INSERT INTO Project_Module(project_module_name, project_id) \
                SELECT TOP 1 '%s', %s FROM Project_Module\
                WHERE NOT EXISTS (SELECT Null FROM Project_Module WHERE project_id=%s AND project_module_name = '%s');" % (module[0],module[1],module[1],module[0])

        # print(sql)
    
        cursor.execute(sql)
        conn.commit()

    print(str(count) + ' project modules inserted')


def insert_status(status_name): # status_name
    count = 0
    sql = "INSERT INTO Resolution_Status (status_name) SELECT TOP 1 '%s' FROM Resolution_Status\
            WHERE NOT EXISTS (SELECT * FROM Resolution_Status WHERE status_name = '%s');" % (status_name,status_name)

    cursor.execute(sql)
    count += 1

    conn.commit()

    print(str(count) + ' status inserted')

def insert_severity(severity_name): # severity_name
    count = 0
    sql = "INSERT INTO Severity (severity_name) SELECT TOP 1 '%s' FROM Severity\
            WHERE NOT EXISTS (SELECT * FROM Severity WHERE severity_name = '%s');" % (severity_name,severity_name)

    cursor.execute(sql)
    count += 1

    conn.commit()

    print(str(count) + ' severity inserted')

def insert_artifact(artifact_name): # artifact_name
    count = 0
    sql = "INSERT INTO Artifact_Type (artifact_name) SELECT TOP 1 '%s' FROM Artifact_Type\
            WHERE NOT EXISTS (SELECT * FROM Artifact_Type WHERE artifact_name = '%s');" % (artifact_name,artifact_name)

    cursor.execute(sql)
    count += 1

    conn.commit()

    print(str(count) + ' severity inserted')

def insert_finding_category(finding_name,artifact_name): # finding_name,artifact_name
    count = 0
    artifact_id = query_artifact_id(artifact_name)
    if(artifact_id == ''):
        print("Something Wrong")

    sql = "INSERT INTO Finding_Category (finding_name,artifact_type_id) SELECT TOP 1 '%s',%s FROM Finding_Category\
            WHERE NOT EXISTS (SELECT * FROM Finding_Category WHERE finding_name = '%s');" % (finding_name,artifact_id,finding_name)

    cursor.execute(sql)
    count += 1

    conn.commit()

    print(str(count) + ' severity inserted')

# def update_observations(observations):

def insert_pa_access(pa_dict):
    print("insert pa access")

    project_module_id = query_project_module_by_name(pa_dict['project_module'])
    artifact_location = pa_dict['artifact_location']
    count = pa_dict['count']
    artifact_type_id = query_artifact_id(pa_dict['artifact_type'])
    doc_name_version = str(pa_dict['doc_name_version'])
    doc_name_version = doc_name_version.replace("\'",'')
    review_by = str(pa_dict['review_by'])
    review_by.rstrip()
    project_id = query_projects_id(pa_dict['project_name'])
    team_id = query_team_id(pa_dict['vendor_name'])
    
    # print(project_module_id,artifact_location,count,artifact_type_id,doc_name_version,review_by,project_id,team_id)

    if project_module_id == '':
        project_module_id = "NULL"

    if review_by == '':
        review_by = "NULL"
    
    if team_id == '':
        team_id = "NULL"

    if doc_name_version == '':
        doc_name_version = "NULL"

    sql = "SELECT TOP 1  PA_id FROM Project_Artifact ORDER BY PA_id DESC;"
    old_pa_id = query_sql(sql)
    # print(old_pa_id)

    sql_project_artifact_insert = "INSERT INTO Project_Artifact (artifact_location, team_id, doc_name_version, [count], review_by, artifact_type_id, project_id, project_module_id )\
    SELECT TOP 1 '%s',%s,'%s',%s,'%s',%s,%s,%s FROM Project_Artifact;" % (artifact_location,team_id,doc_name_version,count,review_by, artifact_type_id, project_id, project_module_id)
    # print(sql_project_artifact_insert)
    cursor.execute(sql_project_artifact_insert)
    conn.commit()


    new_pa = "SELECT TOP 1  PA_id FROM Project_Artifact ORDER BY PA_id DESC;"
    new_pa_id = query_sql(new_pa)
    # print(new_pa_id)

    if new_pa_id == old_pa_id:
        print('Cannot insert into Project Artifact table with below sql:')
        print(sql_project_artifact_insert)
        return -1
    
    pa_id = new_pa_id[0][0]

    return pa_id


def insert_finding(find_dict):
    print("insert finding access")

    pa_id = find_dict['pa_id']
    artifact_type_id = query_artifact_id(find_dict['artifact'])
    finding_category_id = query_finding_category_id(find_dict['finding_category'])
    status_id = query_status_id(find_dict['status'])
    severity_id = query_severity_id(find_dict['severity'])

    location_in_artifact = str(find_dict['location_in_artifact'])
    qa_review_observation = str(find_dict['qa_review_observation'])
    project_response = str(find_dict['project_response'])

    review_date = datetime.datetime.strptime(find_dict['review_date'], '%m/%d/%Y').date()
    followup_1 = datetime.datetime.strptime(find_dict['followup_1'], '%m/%d/%Y').date()
    followup_2 = datetime.datetime.strptime(find_dict['followup_2'], '%m/%d/%Y').date()
    followup_3 = datetime.datetime.strptime(find_dict['followup_3'], '%m/%d/%Y').date()
    followup_4 = datetime.datetime.strptime(find_dict['followup_4'], '%m/%d/%Y').date()
    followup_5 = datetime.datetime.strptime(find_dict['followup_5'], '%m/%d/%Y').date()
    followup_6 = datetime.datetime.strptime(find_dict['followup_6'], '%m/%d/%Y').date()
    followup_7 = datetime.datetime.strptime(find_dict['followup_7'], '%m/%d/%Y').date()
    followup_8 = datetime.datetime.strptime(find_dict['followup_8'], '%m/%d/%Y').date()
    followup_9 = datetime.datetime.strptime(find_dict['followup_9'], '%m/%d/%Y').date()
    followup_10 = datetime.datetime.strptime(find_dict['followup_10'], '%m/%d/%Y').date()

    # print("-----------------\n\n")
    # print(pa_id,artifact_type_id,location_in_artifact,finding_category_id,qa_review_observation,project_response,status_id,severity_id)
    # print(review_date,followup_1,followup_2,followup_3,followup_4,followup_5,followup_6,followup_7,followup_8,followup_9,followup_10)
    # print("-----------------\n\n")

    if qa_review_observation == '':
        qa_review_observation = "NULL"
    
    if project_response == '':
        project_response = "NULL"

    if finding_category_id == '':
        finding_category_id = "NULL"

    # get the top 1 id
    sql_id = "SELECT TOP 1 Findings_id FROM Findings ORDER BY Findings_id DESC;"
    old_id = query_sql(sql_id)

    sql = "INSERT INTO Findings(PA_id,location_in_artifact,finding_category_id,QA_review_observation,project_response,status_id,severity_id,review_date,followup_review_date_1,\
            followup_review_date_2,followup_review_date_3,followup_review_date_4,followup_review_date_5,followup_review_date_6,followup_review_date_7,\
            followup_review_date_8,followup_review_date_9,followup_review_date_10) SELECT TOP 1 %s,'%s', %s, '%s' , '%s' , %s , %s , '%s' , '%s' , '%s' , '%s' ,'%s','%s','%s','%s','%s','%s','%s'  FROM Findings;" % (pa_id,location_in_artifact,finding_category_id,qa_review_observation,project_response,status_id, severity_id,review_date,followup_1,followup_2,followup_3,followup_4,followup_5,followup_6,followup_7,followup_8,followup_9,followup_10)

    # print(sql)
    cursor.execute(sql)
    conn.commit()

    sql_query_finding_id = "SELECT TOP 1 Findings_id FROM Findings ORDER BY Findings_id DESC;"
    new_id = query_sql(sql_query_finding_id)

    # sql_query_finding_id = "SELECT * FROM Findings WHERE Findings_id = %s;" % (int(new_id[0][0]))
    # test = query_sql(sql_query_finding_id)
    # print(test)
    # print(query_result)


    if new_id == old_id:
        print('Cannot insert into Findings table with below sql:')
        print(sql)
        return -1
    
    finding_id = new_id[0][0]
    # print(finding_id)

    return finding_id


'''
    Update in Access

'''

def update_project_artifact(pa_dict):
    print("update pa access")

    count = pa_dict['count']
    doc_name_version = str(pa_dict['doc_name_version'])
    doc_name_version = doc_name_version.replace("\'",'')
    review_by = str(pa_dict['review_by'])
    review_by.rstrip()

    if review_by == '':
        review_by = "NULL"

    if doc_name_version == '':
        doc_name_version = "NULL"

    sql = "UPDATE Project_Artifact \
           SET count = %s,doc_name_version= '%s',review_by='%s' \
           WHERE PA_id = (SELECT PA_id FROM Findings WHERE Findings_id = %s)" % (count,doc_name_version,review_by,pa_dict['finding_id'])

    # print(sql)
    update_sql(sql)


def update_finding(find_dict):

    review_date = find_dict['review_date']
    followup_1 = find_dict['followup_1']
    followup_2 = find_dict['followup_2']
    followup_3 = find_dict['followup_3']
    followup_4 = find_dict['followup_4']
    followup_5 = find_dict['followup_5']
    followup_6 = find_dict['followup_6']
    followup_7 = find_dict['followup_7']
    followup_8 = find_dict['followup_8']
    followup_9 = find_dict['followup_9']
    followup_10 = find_dict['followup_10']

    # print("-----------------\n\n")
    # print(pa_id,artifact_type_id,location_in_artifact,finding_category_id,qa_review_observation,project_response,status_id,severity_id)
    # print(review_date,followup_1,followup_2,followup_3,followup_4,followup_5,followup_6,followup_7,followup_8,followup_9,followup_10)
    sql = "UPDATE Findings AS F, Resolution_Status AS RS, Finding_Category AS FC, Severity AS S \
            SET F.location_in_artifact = '%s', F.finding_category_id = FC.finding_category_id, F.QA_review_observation = '%s', \
            F.project_response = '%s',F.status_id = RS.status_id, F.severity_id = S.severity_id, F.review_date = #%s#,   \
            followup_review_date_1 = #%s#,followup_review_date_2 = #%s#,followup_review_date_3 = #%s#,followup_review_date_4 = #%s#,\
            followup_review_date_5 = #%s#,followup_review_date_6 = #%s#,followup_review_date_7 = #%s#,followup_review_date_8 = #%s#,\
            followup_review_date_9 = #%s#,followup_review_date_10 = #%s#\
            WHERE F.Findings_id = %s AND RS.status_name = '%s' AND S.severity_name = '%s' AND FC.finding_name = '%s';" \
            % (find_dict['location_in_artifact'],find_dict['qa_review_observation'],find_dict['project_response'],find_dict['review_date'],\
                find_dict['followup_1'],find_dict['followup_2'],find_dict['followup_3'],find_dict['followup_4'],find_dict['followup_5'],\
                find_dict['followup_6'],find_dict['followup_7'],find_dict['followup_8'],find_dict['followup_9'],find_dict['followup_10'],\
                find_dict['finding_id'],find_dict['status'],find_dict['severity'],find_dict['finding_category'])
    # print(sql)
    update_sql(sql)


#close the connection
def close_connection():
    conn.close() 

'''
    Utility method
'''
# Execute sql and return the results
def query_sql(sql):
    
    cursor.execute(sql)
    rows = cursor.fetchall()

    return rows

def update_sql(sql):

    cursor.execute(sql)
    conn.commit()