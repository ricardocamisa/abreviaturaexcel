🧩 Função VBA para Abreviação de Nomes no Excel
Este repositório contém uma função personalizada em VBA para Excel que permite abreviar nomes completos de forma inteligente, respeitando um limite de caracteres definido. A função é especialmente útil em sistemas de emissão de documentos, crachás, relatórios ou listas, onde o espaço é limitado.

🔹 Principais recursos:

Abrevia os nomes do meio mantendo o primeiro e último nome intactos.

Respeita preposições como de, da, do, das, dos, preservando-as completas quando necessário.

Permite definir um limite máximo de caracteres para o nome final.

Totalmente personalizável e fácil de integrar a qualquer planilha.

🛠️ Exemplo de uso:

vba
Copy
Edit
=AbreviarNomeLongos("Maria das Dores de Almeida Costa", 26)
' Resultado: Maria das D. A. Costa
