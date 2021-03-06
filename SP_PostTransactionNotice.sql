
ALTER proc [dbo].[SBO_SP_PostTransactionNotice]

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
	IF(@object_type='13' AND @transaction_type IN ('A'))
	BEGIN
		
		IF EXISTS( 
					SELECT T0.[DocEntry]
					FROM [dbo].[OINV] T0
					WHERE T0.[DocSubType]='--'
					AND T0.[DocEntry]=@list_of_cols_val_tab_del
					)
		BEGIN
			INSERT INTO [HC].[dbo].[FacturasMigrar]
			SELECT T0.[DocEntry]
					FROM [dbo].[OINV] T0
					WHERE T0.[DocEntry]=@list_of_cols_val_tab_del
		END
	END


--------------------------------------------------------------------------------------------------------------------------------

-- Select the return values
select @error, @error_message

end
