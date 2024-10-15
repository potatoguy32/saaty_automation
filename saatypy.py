import pandas as pd
import string


def get_saaty_template(origin_path, source_sheet_name, output_name):
    df_origen = pd.read_excel(origin_path, sheet_name=source_sheet_name)
    output_objects = {}

    with pd.ExcelWriter(output_name + ".xlsx") as f:
        df_origen.to_excel(f, sheet_name="source_data", index=False)

        matrix = create_initial_matrix(df_origen, "capa")
        weights_matrix, weights_vector = create_weights_objects(matrix)
        
        matrix.to_excel(f, sheet_name="CapaxCapa", startcol=0)
        weights_matrix.to_excel(f, sheet_name="CapaxCapa", startcol=weights_matrix.shape[1] + 1 + 3)
        weights_vector.to_excel(f, sheet_name="CapaxCapa", startcol=(weights_matrix.shape[1] + 1 + 3) * 2)
        output_objects["CapaxCapa"] = (matrix, weights_matrix, weights_vector)
        
        available_capas = df_origen.capa.unique()
        for capa in available_capas:
            current_df = df_origen.query("capa == @capa").copy()
            matrix = create_initial_matrix(current_df, "variable")
            weights_matrix, weights_vector = create_weights_objects(matrix)
            
            matrix.to_excel(f, sheet_name=capa, startcol=0)
            weights_matrix.to_excel(f, sheet_name=capa, startcol=weights_matrix.shape[1] + 1 + 3)
            weights_vector.to_excel(f, sheet_name=capa, startcol=(weights_matrix.shape[1] + 1 + 3) * 2)
            output_objects[capa] = (matrix, weights_matrix, weights_vector)

    return output_objects

def create_initial_matrix(df, reference_column):
    vertical_axis = df.get([reference_column]).drop_duplicates()
    horizontal_axis =  vertical_axis.copy().set_index(reference_column).transpose()
    matrix = pd.concat([vertical_axis, horizontal_axis], axis=0, ignore_index=True).set_index(reference_column)
    matrix[matrix.columns] = None
    row_number, col_number = matrix.shape
    matrix = pd.concat([matrix, matrix.iloc[[row_number - 1, ], :]], axis=0)
    current_index = list(matrix.index)
    current_index[-1] = 'Suma'
    matrix.index = current_index
    for i in range(row_number):
        for j in range(col_number):
            if i == j:
                matrix.iloc[i, j] = 1
            elif i < j:
                matrix.iloc[i, j] = '= 1/{col}{row}'.format(col=get_excel_column(i + 1), row=j + 2)
        
        matrix.iloc[j + 1, i] = '= SUM({start}:{end})'.format(start=get_excel_column(i + 1) + "2", end=get_excel_column(i+1) + str(row_number + 1))
    
    return matrix


def create_weights_objects(matrix):
    weights_matrix = matrix.iloc[:matrix.shape[0] - 1, :].copy()
    weights_matrix[weights_matrix.columns] = ""
    row_num, col_num = weights_matrix.shape[0], weights_matrix.shape[1]
    weights_matrix["PONDERACION"] = ""
    weights_vector = weights_matrix["PONDERACION"].copy()

    for i in range(row_num):
        for j in range(col_num):
            weights_matrix.iloc[i, j] = "= {numerator}/{denominator}".format(numerator=get_excel_column(j + 1) + str(i + 2), denominator=get_excel_column(j + 1) + str(row_num + 2))
        
        weights_matrix.iloc[i, j + 1] = "= SUM({start}:{end})".format(start=get_excel_column(row_num + 3 + 3) + str(i + 2), end=get_excel_column(row_num + row_num + 3 + 2) + str(i + 2))
        weights_vector.iloc[i] = "= {position}/{length}".format(position=get_excel_column(row_num + row_num + 3 + 3) + str(i + 2), length=col_num)

    weights_vector = pd.DataFrame(weights_vector)
    
    return weights_matrix, weights_vector


def get_excel_column(num):
    if num <= 25:
        return string.ascii_uppercase[num]
    
    letter = ''
    while num > 25:   
        letter += chr(65 + int((num) / 26) - 1)
        num = num - (int((num)/26)) * 26

    letter += chr(65 + (int(num)))

    return letter
