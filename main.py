
# Importacoes e Configuracao do Pandas 
import pandas as pd

pd.set_option('display.width', 1000)
pd.set_option('display.max_columns', None)


# Função para Leitura  

def read_file_excel(file_path, column_names, skiprows=1):

    try:
        return pd.read_excel(file_path, skiprows=skiprows, names=column_names)
    
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {file_path}")
        return None
    except pd.errors.EmptyDataError:
        print(f"Arquivo vazio: {file_path}")
        return None
    except pd.errors.ParserError:
        print(f"Erro ao analisar o arquivo: {file_path}")
        return None
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao ler o arquivo {file_path}: {e}")
        return None

def write_file_excel(file_path, df, index=False):
    try:
        df.to_excel(file_path, index=index)
        print(f"Arquivo {file_path} escrito com sucesso.")

    except PermissionError:
        print(f"Permissão negada ao tentar escrever o arquivo {file_path}. Verifique se o arquivo está aberto ou se você tem permissão para escrever no diretório.")
    
    except FileNotFoundError:
        print(f"Caminho do arquivo não encontrado: {file_path}. Verifique se o diretório existe.")
    
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao escrever o arquivo {file_path}: {e}")


# Variaveis 
PATH_INPUT_SHEET_ORDERS_RELEASED = "./data_input/Planilha_Azul.xlsx"
PATH_INPUT_SHEET_ORDERS_BLOCKED = "./data_input/Planilha_Amarela.xlsx"

COLUMN_NAME_SHEET_ORDERS_RELEASED = ["Confirmed_receipt_date","Review_date","Mark","Release_reason","Sales_order","Name","Sales_total","Customer_account","Warehouse","Customer_group","Released","Sales_responsible","Credit_control_group","Modified_by","Active","Released_by","Sales_taker","Credit_control_number","Masterpack_reference","Ship_date","Credit_control_reason","Document_status","Load_ID","Terms_of_payment","Forced_hold_reason","Company"
]
COLUMN_NAME_SHEET_ORDERS_BLOCKED = ["On_First_Failed_Queue","KPI","Before_9am","Advanced","Morning_Queue","Error_Type","Replen_Line","Cost_Centre","Warehouse","Confirmed_pick_date","Ship_complete2","Do_not_consolidate","In_credit_control","Expedite","Delivery_zone","Sales_order","Sales_origin","Customer","Name","Item_number","Customer_reference","Quantity_available_to_release","Quantity","Product_name","Inventory_unit","Site","Batch_Number","Location","Inventory_status","Licence_plate","Delivery_name","Customer_reference2","City","State","Postcode","Confirmed_receipt_date","Modified_by","Mode_of_delivery","Modified_date_and_time","Created_date_and_time","Address","Do_not_process","Sales_Group","Customer_group","Product_posting_group","COGs_Estimate"]

COLUMN_NAME_OUTPUT_FILE = []
OUTPUT_FILE = "./data_output/ORDERS_BLOCKED.xlsx"



# Leitura das Planilhas 
df_orders_released = read_file_excel(PATH_INPUT_SHEET_ORDERS_RELEASED, column_names=COLUMN_NAME_SHEET_ORDERS_RELEASED ,skiprows=1)
df_orders_blocked  = read_file_excel(PATH_INPUT_SHEET_ORDERS_BLOCKED,  column_names=COLUMN_NAME_SHEET_ORDERS_BLOCKED  ,skiprows=2)

## Remoção de Duplicados - Regra 1 
df_orders_blocked = df_orders_blocked.drop_duplicates(subset=["Sales_order"]) # Remover duplicados apenas dos pedidos bloqueados

## Filtrar a coluna Credit Control - Regra 2
df_orders_blocked = df_orders_blocked[df_orders_blocked["In_credit_control"] == "Yes"]

## Filtrar a coluna Cost Center por Residencial - Regra 3 
df_orders_blocked = df_orders_blocked[df_orders_blocked["Cost_Centre"].str.contains("Residential")]
 

# Merge entre as duas planilhas utilizando o Sales_Order como chave. 
df_merge_orders = pd.merge(df_orders_released, df_orders_blocked, on="Sales_order", how="inner")

quantidade_linhas = df_merge_orders.shape[0]

if quantidade_linhas != 0:
    #df_merge_orders = df_merge_orders[COLUMN_NAME_OUTPUT_FILE]  # Caso necessario colunas especificas.
    write_file_excel(OUTPUT_FILE, df_merge_orders, index=False)




