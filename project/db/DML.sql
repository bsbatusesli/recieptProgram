USE nokia;


INSERT INTO saved_items(itemCode,shortDesc,longDesc,price) VALUES ('3HE10493AB', 'SYS - 7210 SAS-Sx 22F2C4SFP+', 'SYS - 7210 SAS-Sx 22F2C4SFP+', 4600);
INSERT INTO saved_items(itemCode,shortDesc,longDesc,price) VALUES ('3HE10498AA', 'PS - 7210 SAS-Sx Power Supply AC (F)', 'PS - 7210 SAS-Sx Power Supply (fits fiber 1/10G Sx units only) AC 400W', 50);
INSERT INTO saved_items(itemCode,shortDesc,longDesc,price) VALUES ('3HE09326AA', 'SFP+10GE SR - LC ROHS6/6-40/85C', '1-port 10GBASE-SR Small Form-Factor Pluggable+ (SFP+) Optics Module, 850 nm, 26 to 300 meters, LC Connector, RoHS 6/6 compliant, extended temperature range -40/85C', 380.07);

INSERT INTO switches(iid, noofport, noofuplink, wirespeed) VALUES (1, 24, 2, 128);
INSERT INTO optics(iid, rangelimit,bandwidth) VALUES (3, 300, 10);

INSERT INTO bundles(name) VALUES ('bundle 1');
INSERT INTO bundles(name) VALUES ('bundle 2');

INSERT INTO bundle_items(bid, iid) VALUES (1, 1, 5);
INSERT INTO bundle_items(bid, iid) VALUES (1, 2, 2);
INSERT INTO bundle_items(bid, iid) VALUES (2, 1, 1);
INSERT INTO bundle_items(bid, iid) VALUES (2, 3, 2);
