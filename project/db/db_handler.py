
import mysql.connector
from mysql.connector import Error
from src.Bundle import *

class db():

    def __init__(self, _host,_database,_user,_password):
        self.host = _host
        self.database = _database
        self.user = _user
        self.password = _password

    def get_connection(self):


        connection = mysql.connector.connect(host=self.host,
                                            database=self.database,
                                            user=self.user,
                                            password=self.password)
        return connection




    def close_connection(self,connection):
        if connection:
            connection.close()




    def getAllSavedParts(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """SELECT * FROM saved_items"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records

        except Error as e:
            print("Error while getting data", e)

    def getAllSwitches(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT si.iid,itemCode,shortDesc,longDesc,price,noofport,noofuplink,wirespeed,throughtput FROM saved_items AS si
                                INNER JOIN switches ON si.iid = switches.iid"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records
        except Error as e:
            print("Error while getting data", e)

    def getAllOptics(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT si.iid,itemCode,shortDesc,longDesc,price,rangelimit,bandwidth FROM saved_items AS si
                                INNER JOIN optics ON si.iid = optics.iid"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records
        except Error as e:
            print("Error while getting data", e)

    def getOtherParts(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """SELECT * FROM saved_items WHERE iid != (SELECT iid FROM switches ) AND iid != (SELECT iid FROM optics )"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records
        except Error as e:
            print("Error while getting data", e)

    def getAllBundles(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """SELECT * FROM bundles"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records

        except Error as e:
            print("Error while getting data", e)
    
    def getAllBundleNames(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """SELECT name FROM bundles"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            names = []
            for row in records:
                names.append(row[0]) 
            return names

        except Error as e:
            print("Error while getting data", e)

    def getBundleItems(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT b.bid, name, bi.iid, quantity  FROM bundles AS b
                                INNER JOIN bundle_items AS bi ON b.bid = bi.bid
                                ORDER BY bid"""
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return records

        except Error as e:
            print("Error while getting data", e)

    def getPartInfo(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT * FROM saved_items 
                                WHERE itemCode = %s"""
            cursor.execute(select_query, (itemCode,))
            records = cursor.fetchall()
            self.close_connection(connection)
            return records

        except Error as e:
            print("Error while getting data", e)

    def getPartCount(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT COUNT(iid) FROM saved_items """
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return (records[0][0])

        except Error as e:
            print("Error while getting data", e)
    
    def getBundleCount(self):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT COUNT(bid) FROM bundles """
            cursor.execute(select_query)
            records = cursor.fetchall()
            self.close_connection(connection)
            return (records[0][0])

        except Error as e:
            print("Error while getting data", e)
    
    def updatePart(self,_itemCode, _shortDesc, _longDesc, _price, _iid):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            update_query = """  UPDATE saved_items
                                SET itemCode =%s,shortDesc =%s,longDesc =%s,price =%s  WHERE iid = %s """
            cursor.execute(update_query, (_itemCode, _shortDesc, _longDesc, int(_price), _iid,))
            connection.commit()
            self.close_connection(connection)
            return True

        except Error as e:
            print("Error while updating data", e)

    def insertPart(self,_itemCode, _shortDesc, _longDesc, _price):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            insert_query = """  INSERT INTO saved_items (itemCode,shortDesc,longDesc,price)
                                VALUES (%s,%s,%s,%s) """
            cursor.execute(insert_query, (_itemCode, _shortDesc, _longDesc, int(_price),))
            connection.commit()
            self.close_connection(connection)
            return True

        except Error as e:
            print("Error while updating data", e)

    def updateSwitch(self,_noofport,_noofuplink,_wirespeed,_throughtput,_iid):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            update_query = """  UPDATE switches
                                SET noofport =%s,noofuplink =%s,wirespeed=%s,throughtput =%s  WHERE iid = %s """
            cursor.execute(update_query, (_noofport, _noofuplink, _wirespeed, _throughtput, _iid,))
            connection.commit()
            self.close_connection(connection)
            return True
        except Error as e:
            print("Error while updating data", e)

    
    def updateOptics(self,_rangelimit,_bandwidth,_iid):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            update_query = """  UPDATE optics
                                SET rangelimit =%s,bandwidth =%s WHERE iid = %s """
            cursor.execute(update_query, (int(_rangelimit), int(_bandwidth), _iid,))
            connection.commit()
            self.close_connection(connection)
            return True

        except Error as e:
            print("Error while updating data", e)

    def insertBundle(self,_bundle, _bid,_partList):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            insert_query = """  INSERT INTO bundles (bid,name)
                                    VALUES (%s,%s) """
            cursor.execute(insert_query, (_bid, _bundle.getName(),))
            connection.commit()
            parts = _bundle.getConnectedParts()
            for part in parts:
                insert_query = """  INSERT INTO bundle_items (bid,iid,quantity)
                                    VALUES (%s,%s,%s) """   
                index = _partList.index(part[0])
                cursor.execute(insert_query, (_bid, index,part[1] ,))
                connection.commit()
            self.close_connection(connection)
            return True

        except Error as e:
            print("Error while updating data", e)

    def updateBundle(self,_bundle, _bid,_partList):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            parts = _bundle.getConnectedParts()
            for part in parts:
                index = _partList.index(part[0])
                if self.isPartExistsInBundle(_bid,index):
                    update_query = """  UPDATE bundle_items
                                    SET bid =%s,iid =%s ,quantity=%s WHERE bid = %s AND iid = %s """
                    cursor.execute(update_query, (_bid, index, part[1],_bid, index))
                    connection.commit()
                else:
                    insert_query = """  INSERT INTO bundle_items (bid,iid,quantity)
                                    VALUES (%s,%s,%s) """                 
                    index = _partList.index(part[0])
                    cursor.execute(insert_query, (_bid, index,part[1] ,))
                    connection.commit()
            self.close_connection(connection)
            return True

        except Error as e:
            print("Error while updating data", e)
    
    def isPartExistsInBundle(self,_bid,_iid):
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            select_query = """  SELECT COUNT(*) FROM bundle_items WHERE bid =%s AND iid =%s """
            cursor.execute(select_query,(_bid,_iid,))
            records = cursor.fetchall()
            self.close_connection(connection)
            if records[0][0] == 0:
                return False
            else:
                return True

        except Error as e:
            print("Error while getting data", e)
