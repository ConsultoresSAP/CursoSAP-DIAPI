ALTER proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(30), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int, 
@list_of_key_cols_tab_del nvarchar(255), 
@list_of_cols_val_tab_del nvarchar(255) 

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

/*Error 1300001: Validación de comentario vacio en factura */	
	IF(@object_type='13' AND @transaction_type IN ('A','U'))
	BEGIN
		
		IF EXISTS( 
					SELECT T0.[DocEntry]
					FROM [dbo].[OINV] T0
					WHERE T0.[DocSubType]='--'
					AND ISNULL(T0.[Comments],'')=''
					AND T0.[DocEntry]=@list_of_cols_val_tab_del

					)
		BEGIN
			SET @error=1300001
			SET @error_message = 'No puede quedar el comentario vacio'
		END
	END
/* FIN ERROR 1300001 */

/*Error 2400001: Cancelar Facturas al contado */	
	IF(@object_type='24' AND @transaction_type IN ('A'))
	BEGIN
		
		IF EXISTS( 
					SELECT T0.[DocEntry]
					FROM [dbo].[ORCT] T0
					WHERE 
					ISNULL(T0.[U_Caja],'')=''
					AND T0.[DocEntry]=@list_of_cols_val_tab_del

					)
		BEGIN
			SET @error=2400001
			SET @error_message = 'Pago cancelado'
		END
	END
/* FIN ERROR 2400001 */

/*Error 200001: Validación de inactivar SN con saldo */	
	IF(@object_type='2' AND @transaction_type IN ('U'))
	BEGIN
		
		IF EXISTS( 
					SELECT T0.[CardCode]
					FROM [dbo].[OCRD] T0
					WHERE 
					T0.[CardCode]=@list_of_cols_val_tab_del
					AND ISNULL(T0.[Balance],0) != 0
					AND T0.[frozenFor]='Y'
					)
		BEGIN
			SET @error=200001
			SET @error_message = 'Error: No debe inactivarse un SN con saldo'
		END
	END
/* FIN ERROR 200001 */

/*Error 1300001: Validación de editar estados de Honduras */	
	IF(@object_type='130' AND @transaction_type IN ('U'))
	BEGIN
		
		IF ( LEFT(@list_of_cols_val_tab_del,CHARINDEX('	',@list_of_cols_val_tab_del)-1) ='HN' )
		--SUBSTRING(@list_of_cols_val_tab_del,CHARINDEX('	',@list_of_cols_val_tab_del)+1,len(@list_of_cols_val_tab_del)
		BEGIN
			SET @error=1300001
			SET @error_message = 'Error: No puede editar estados de Honduras '+@list_of_cols_val_tab_del
		END
	END
/* FIN ERROR 1300001 */

/*Error 1120001: Borrador de traslado de inventario */	
	IF(@object_type='112' AND @transaction_type IN ('A'))
	BEGIN
		
		IF exists( 
					SELECT T0.[DocEntry] 
					FROM [dbo].[ODRF] T0
					WHERE
					T0.[DocEntry]=@list_of_cols_val_tab_del
					AND T0.[ObjType]='67'
					AND ISNULL(T0.[Comments],'')=''
					)
		BEGIN
			SET @error=1120001
			SET @error_message = 'Error: puede quedar el comentario vacio '
		END
	END
/* FIN ERROR 1120001 */

	

-- Select the return values
select @error, @error_message

end
