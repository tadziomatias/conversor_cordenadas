import pandas as pd

# Função para converter coordenadas de graus, minutos e segundos para graus decimais
def dms_to_dd(coord):
    # Substituir o símbolo de segundo e direção oeste
    coord = coord.replace('”', '"').replace('O', 'W')
    
    # Verifica se a coordenada está no formato esperado
    if '°' not in coord or '\'' not in coord or '\"' not in coord:
        raise ValueError("Formato de coordenadas inválido")

    # Divide a string de coordenadas
    degrees, rest = coord.split('°')
    minutes, rest = rest.split('\'')
    seconds, direction = rest.split('"')

    # Converte as partes em números inteiros
    degrees = int(degrees)
    minutes = int(minutes)
    seconds = float(seconds)

    # Calcula os graus decimais
    dd = degrees + minutes / 60 + seconds / 3600

    # Verifica a direção (N, S, E, W) e ajusta o sinal, se necessário
    if direction in ('S', 'W'):
        dd *= -1

    return dd

# Ler os dados da planilha Excel
df = pd.read_excel('/Users/matias/Documents/codigo/pontos.xlsx')

# Exibir as primeiras linhas do DataFrame para garantir que os dados foram lidos corretamente
print(df.head())

# Converter as coordenadas para graus decimais
df['Lat'] = df['Lat'].apply(dms_to_dd)
df['Long'] = df['Long'].apply(dms_to_dd)

# Exibir as primeiras linhas do DataFrame para garantir que a conversão foi feita corretamente
print(df.head())

# Salvar o DataFrame com as coordenadas convertidas em um novo arquivo Excel
output_path = '/Users/matias/Documents/codigo/pontos_convertidos.xlsx'
df.to_excel(output_path, index=False)

print(f"Conversão concluída e arquivo salvo como '{output_path}'.")
