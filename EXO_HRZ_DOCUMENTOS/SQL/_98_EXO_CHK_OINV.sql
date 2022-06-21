CREATE PROCEDURE _98_EXO_CHK_OINV
(
	IN pObject_type NVARCHAR(30),
	IN pTransaction_type NCHAR(1),
	IN pList_of_cols_val_tab_del NVARCHAR(255),
	OUT pError INT,
	OUT pError_message NVARCHAR(200)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
vPAGO NVARCHAR(100);
	
BEGIN		
	IF :pObject_type = '13' THEN
	IF (:pTransaction_type = 'A' OR :pTransaction_type = 'U') THEN
			SELECT T1."Descript" INTO vPAGO
			FROM OINV T0  INNER JOIN OPYM T1 ON T0."PeyMethod" = T1."PayMethCod"
			WHERE t0."DocEntry" = :pList_of_cols_val_tab_del;
					
			IF :vPAGO <> '' THEN
				DECLARE EXIT HANDLER FOR SQLEXCEPTION
					BEGIN
						IF ::SQL_ERROR_CODE <> 0 THEN
							pError := 1;
							pError_message := '(EXO) ' || ::SQL_ERROR_CODE || ' ' || ::SQL_ERROR_MESSAGE;
						END IF;
					END;
				
				UPDATE "OINV" SET "U_EXO_VPAGO" = vPAGO
				WHERE "DocEntry" = :pList_of_cols_val_tab_del;
			END IF;						
		END IF;
	END IF;
END;