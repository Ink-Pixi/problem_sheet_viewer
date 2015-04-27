import os
import pypyodbc

class DataBase(object):
    def __init__(self):
        super(object, self).__init__()
    
    def connect_ps(self):
        #problem sheet connection.
        conn = pypyodbc.connect('DRIVER={SQL Server}; SERVER=SQLSERVER; DATABASE=ProblemSheets; Trusted_Connection=yes')
        db = conn.cursor()
        
        return db

    def get_table_data(self, tblType):
        db = self.connect_ps()
        sql = """
                SELECT
                ps.ID,
                cc.companyDesc,
                CASE 
                    WHEN (select COUNT(ID) from dbo.tblChangesAdditions where problemSheetID = ps.ID) >= 1 THEN 'Changes Additions'
                    WHEN (select COUNT(ID) from dbo.tblMisShip where problemSheetID = ps.ID) >= 1 THEN 'Mis Ship'
                    WHEN (select COUNT(ID) from dbo.tblDefective where problemSheetID = ps.ID) >= 1 THEN 'Defective'
                    WHEN (select COUNT(ID) from dbo.tblExpedited where problemSheetID = ps.ID) >= 1 THEN 'Expedited'
                    ELSE 'N/A'
                END AS Type,
                CASE 
                    WHEN cc.ID = 3 
                    THEN ps.invoiceNumber 
                    ELSE ps.shortNumber
                END AS shortNumber,"""
                
        if tblType == 0:
            sqlType = """
                            
                CONVERT(VARCHAR(12), ps.callDate, 101) callDate,
                CASE ps.isPriority
                    WHEN 1 THEN 'Yes'
                    WHEN 0 THEN NULL
                END AS priorityText,

                CASE
                    WHEN ps.isComplete = 1 THEN 'Completed'
                    WHEN ps.inProgress = 1 THEN 'In Progress'
                    ELSE NULL
                END AS completeText"""
                
        elif tblType == 1:
            sqlType = """
                CONVERT(VARCHAR(12), ps.dateCompleted, 101) dateCompleted,
                CASE ps.isPriority
                    WHEN 1 THEN 'Yes'
                    WHEN 0 THEN NULL
                END AS priorityText,            
                completedBy"""
                
        sqlEnd = """
            FROM
                dbo.tblProblemSheet ps
            JOIN
                dbo.tblCompany cc ON ps.companyID = cc.ID
            WHERE 
                ps.isComplete = ?        
            ORDER BY priorityText DESC, entryDate ASC """
            
        sql = sql + sqlType + sqlEnd
        
        db.execute(sql, [tblType])
        ds = db.fetchall()
        db.close()
        
        return ds  
    
    def check_progress(self, ID):
        db = self.connect_ps()
        db.execute("SELECT inProgress FROM dbo.tblProblemSheet WHERE ID = ?", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds[0]
    
    def check_complete(self, ID):
        db= self.connect_ps()
        db.execute("SELECT isComplete FROM dbo.tblProblemSheet WHERE ID = ?", [ID])
        ds = db.fetchone()
        
        return ds[0]

    def mark_in_progress(self, ID):
        db = self.connect_ps()
        db.execute("UPDATE dbo.tblProblemSheet SET inProgress = 1, isComplete = 0 WHERE ID = ?", [ID])
        db.commit()        
    
    def mark_complete(self, ID):
        
        db = self.connect_ps()
        db.execute("UPDATE dbo.tblProblemSheet SET inProgress = 0, isComplete = 1 WHERE ID = ?", [ID])
        db.commit()
        db.execute("UPDATE dbo.tblProblemSheet SET dateCompleted = GETDATE() WHERE ID = ?", [ID])
        db.commit()
        db.execute("UPDATE dbo.tblProblemSheet SET completedBy = ? WHERE ID = ?", [os.getlogin(), ID])
        db.commit()
        db.close()
                
    def mark_open(self, ID):
        db = self.connect_ps()
        db.execute("UPDATE dbo.tblProblemSheet SET inProgress = 0, isComplete = 0, completedBy = NULL, dateCompleted = NULL WHERE ID = ?", [ID])
        db.commit()
        db.close()
        
    def check_trace(self, ID):
        db = self.connect_ps()
        db.execute("SELECT isTrace FROM dbo.tblProblemSheet WHERE ID = ?", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds[0]
    
    def check_priority(self, ID):
        db = self.connect_ps()
        db.execute("SELECT isPriority FROM dbo.tblProblemSheet WHERE ID = ?", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds[0]
    
    def get_report_type(self, ID):
        db = self.connect_ps()
        db.execute("""
            SELECT 
                c.companyDesc
            FROM 
                dbo.tblProblemSheet ps
            JOIN
                dbo.tblCompany c ON ps.companyID = c.ID
            WHERE ps.ID = ?""", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds[0]
    
    def check_expedited_approval(self, ID):
        db = self.connect_ps()
        db.execute("SELECT approved FROM dbo.tblExpedited WHERE problemSheetID = ?", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds[0]
    
    def check_credit_info(self, ID):
        db = self.connect_ps()
        db.execute("SELECT problemSheetID FROM dbo.tblCreditCard WHERE problemSheetID = ?", [ID])
        ds = db.fetchone()
        db.close()
        
        return ds
    
    def get_sheet_id(self, searchNumber):
        txtNumber = str(searchNumber)
        db = self.connect_ps()
        db.execute("""
            SELECT 
                ID
            FROM 
                dbo.tblProblemSheet 
            WHERE
                REPLACE(longNumber, '-', '') = ?
            OR
                shortNumber = ?
            OR
                CAST(ID AS VARCHAR(25)) = ?""", [txtNumber, txtNumber, txtNumber])
        ds = db.fetchone()
        db.close()
        
        return ds[0]