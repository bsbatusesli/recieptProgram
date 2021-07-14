USE nokia;


DROP TABLE IF EXISTS switches;
DROP TABLE IF EXISTS optics;
DROP TABLE IF EXISTS bundle_items;
DROP TABLE IF EXISTS saved_items;
DROP TABLE IF EXISTS bundles;

CREATE TABLE saved_items(
	iid int NOT NULL AUTO_INCREMENT,
    itemCode varchar(15) NOT NULL,
    shortDesc varchar(100) NOT NULL,
    longDesc varchar(500) NOT NULL,
    price double NOT NULL,
    PRIMARY KEY (iid) );
    
CREATE TABLE switches(
	iid int NOT NULL,
    noofport int,
    noofuplink int,
    wirespeed int,
    throughtput int,
    PRIMARY KEY(iid),
    FOREIGN KEY (iid) REFERENCES saved_items(iid));
    
CREATE TABLE optics(
	iid int NOT NULL,
    rangelimit int,
    bandwidth int,
    PRIMARY KEY(iid),
    FOREIGN KEY(iid) REFERENCES saved_items(iid));

CREATE TABLE bundles(
	bid int NOT NULL AUTO_INCREMENT,
    name varchar(30) NOT NULL,
    PRIMARY KEY(bid));
    
CREATE TABLE bundle_items(
	iid int NOT NULL,
    bid int NOT NULL,
    quantity int NOT NULL,
    PRIMARY KEY(iid,bid),
    FOREIGN KEY(iid) REFERENCES saved_items(iid),
    FOREIGN KEY(bid) REFERENCES bundles(bid));