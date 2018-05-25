#!/usr/bin/env python3.6

###################################
#     IMPORTACAO  DE LIBRARIES
###################################
import xlsxwriter
import datetime
import calendar
import time

###################################
#     CRIA A DATA DO ARQUIVO
###################################
curdateout = datetime.datetime.today().strftime('%y%m%d')

###################################
#    CRIANDO A PLANILHA EXCEL
###################################
( 'ctm_jobs_'+curdateout+'.csv', 'a' )
workbook = xlsxwriter.Workbook('Bitacora'+curdateout+'.xlsx')

###################################
#   RENOMEIA ABA DO MES CORRENTE
###################################
now = (datetime.datetime.now())
year = (now.year)
#print ("MES : {}".format(now.month))
calmon = now.month - 1
#print ("MES : {}".format(calmon))
nmonths=["Janeiro","Fevereiro","Marco","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
curmonth = (nmonths[calmon])

worksheet = workbook.add_worksheet(curmonth+' '+str(year))
worksheet.hide_gridlines(2)
worksheet.set_zoom(80)

###################################
# FORMATANDO O TAMANHO DA COLUNA
###################################
worksheet.set_column('A:A', 1)
worksheet.set_column('B:B', 7)
worksheet.set_column('C:C', 73)
worksheet.set_column('D:D', 8)
worksheet.set_column('E:E', 8)
worksheet.set_column('F:F', 10)
worksheet.set_column('G:ZZ', 6)

###################################
#  FORMATANDO O TAMANHO DA LINHA
###################################
worksheet.set_row(0, 55)
worksheet.set_row(1, 17)

###################################
# DADOS DO CABECALHO E APLIC FORMAT
###################################
merge_prim = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'font_name': 'Calibri',
    'font_size': 14,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white'
})
worksheet.merge_range('B1:C2', 'Acompanhamento Diario Processamento Batch - Producao', merge_prim)

merge_fixos = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FFDAB9'
    #'fg_color': '#D7E4BC'
})
worksheet.merge_range('D1:D2', 'Horario Inicio', merge_fixos)
worksheet.merge_range('E1:E2', 'Meta Negocio', merge_fixos)
worksheet.merge_range('F1:F2', 'Max Elapsed Time', merge_fixos)


###################################
#        DADOS DA LEGENDA
###################################
leg_prim = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'font_name': 'Calibri',
    'font_size': 14,
    'border': 2,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white'
})

leg_cinza = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'fg_color': '#C0C0C0'
})

leg_azcla = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'fg_color': '#ADD8E6'
})

leg_azesc = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'fg_color': '#4682B4'
})

leg_verde = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'fg_color': 'green'
})

leg_verme = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'fg_color': 'red'
})

vazio=''
worksheet.merge_range('B78:D78', ' *********** L E G E N D A *********** ', leg_prim)
worksheet.write(79,1,vazio,leg_cinza)
worksheet.write(79,2,'Job nao executa nesta data')

worksheet.write(81,1,vazio,leg_azcla)
worksheet.write(81,2,'Job finalizado dentro do Programado (Atendimento da meta em mais de uma hora.)')

worksheet.write(83,1,vazio,leg_azesc)
worksheet.write(83,2,'Job finalizado dentro do Programado (Atendimento da meta em ate uma hora.)')

worksheet.write(85,1,vazio,leg_verde)
worksheet.write(85,2,'Job nao executado devido a Mudanca ou Solicitacao de Servico.')

worksheet.write(87,1,vazio,leg_verme)
worksheet.write(87,2,'Job fora da Meta de Horario de Entrega.')

###################################
#  FORMATANDO OS GRUPOS DOS JOBS
###################################
merge_canto1 = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'rotation': 90,
    'valign': 'vcenter',
    'fg_color': '#F0E68C'
})

merge_canto2 = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'rotation': 90,
    'valign': 'vcenter',
    'fg_color': '#C0C0C0'
})



worksheet.merge_range('B3:B8', 'SCC', merge_canto1)
worksheet.merge_range('B9:B26', 'PCR', merge_canto2)
worksheet.merge_range('B27:B38', 'C3M', merge_canto1)
worksheet.merge_range('B39:B47', 'CTC', merge_canto2)
worksheet.merge_range('B48:B58', 'SCG', merge_canto1)
worksheet.merge_range('B59:B62', 'STD', merge_canto2)
worksheet.merge_range('B63:B64', 'CQL', merge_canto1)
worksheet.merge_range('B65:B75', 'SLC', merge_canto2)

###################################
#  CRIANDO OS DADOS DA PLANILHA
###################################
producao = (
['CIP_SCC_SEND_VARREDURA_ASCC010','00:30','05:59','05:29:00'],
['CIP_SCC_SEND_VARREDURA_SERVMARG','00:45','05:59','01:11:00'],
['CIP_SCC_SEND_VARREDURA_ASCC029','00:47','05:59','01:11:00'],
['CIP_SCC_SEND_VARREDURA_ASCC002','22:13','05:59','01:11:00'],
['CIP_SCC_SEND_VARREDURA_ASCC032','22:00','05:59','07:59:00'],
['CIP_SCC_GERA_DATA_REFERENCIA','04:00','05:59','01:59:00'],
['CIP_NPC_GERAR_DATA_REFERENCIA','05:50','06:00','00:10:00'],
['CIP_NPC_ENVIA1_GRADE_PROC_402','05:50','06:00','00:10:00'],
['CIP_NPC_VARRED_DDA0400_INFORMA_DTREF','05:50','06:00','00:10:00'],
['CIP_NPC_SEND_VARREDURA_ADDA120','06:03','10:00','03:57:00'],
['CIP_NPC_SEND_VARREDURA_ADDA117','07:00','10:00','03:00:00'],
['CIP_NPC_TARIFACAO','16:00','18:00','02:00:00'],
['CIP_NPC_GERAR_ARQUIVO_RCO','06:10','09:30','03:20:00'],
['CIP_NPC_ENVIA_ARQUIVO_RCO','06:10','10:00','03:50:00'],
['CIP_NPC_SEND_VARREDURA_CDDA504_14h00','14:00','18:00','04:00:00'],
['CIP_NPC_SEND_VARREDURA_ADDA504_16h00','16:00','20:00','04:00:00'],
['CIP_NPC_INTEGRACAO_SIFAT_ENVIA_ADDAFAT','16:00','20:00','04:00:00'],
['CIP_NPC_SEND_VARREDURA_ADDA003(Ciclico das 06h30)','06:30','09:00','02:30:00'],
['CIP_NPC_SEND_VARREDURA_ADDA003(Ciclico das 17h)','17:00','20:00','03:00:00'],
['CIP_NPC_SEND_VARREDURA_ADDA003(Ciclico das 20h)','20:00','23:00','03:00:00'],
['CIP_NPC_SEND_VARREDURA_ADDA003(Ciclico das 23h)','23:00','02:00','03:00:00'],
['CIP_NPC_SEND_VARREDURA_ADDA200_DOM','18:00','23:00','05:00:00'],
['CIP_NPC_DTREF_CALCULO','22:00','22:10','00:10:00'],
['CIP_NPC_SEND_VARREDURA_ADDA200_22h30m','22:30','23:30','01:00:00'],
['CIP_C3M_GERAR_DATA_REFERENCIA','05:00','05:59','00:59:00'],
['CIP_C3M_SEND_VARREDURA_ACCC038','04:00','05:49','01:49:00'],
['CIP_C3M_SEND_VARREDURA_ACCC301_16h','16:00','19:59','03:59:00'],
['CIP_C3M_SEND_VARREDURA_ACCC304_16h','16:00','19:59','03:59:00'],
['CIP_C3M_SEND_VARREDURA_CANCOPS','20:00','05:49','09:49:00'],
['CIP_C3M_ENVIA_VARREDURA_ACCC800','21:00','05:49','08:49:00'],
['CIP_C3M_SEND_VARREDURA_ACCC301_21h','21:00','05:49','08:49:00'],
['CIP_C3M_SEND_VARREDURA_ACCC304_21h','21:00','05:49','08:49:00'],
['CIP_C3M_SEND_VARREDURA_ACCC306_21h','21:00','05:49','08:49:00'],
['CIP_C3M_SEND_VARREDURA_ACCC801_21h','21:00','05:49','01:11:00'],
['CIP_C3M_SEND_VARREDURA_ACCC801','22:00','05:49','07:49:00'],
['CIP_C3M_SEND_VARREDURA_V_EXPOOFP','23:30','05:59','06:29:00'],
['CIP_CTC_GERAR_DATA_REFERENCIA','04:00','04:59','00:59:00'],
['CIP_CTC_SEND_VARREDURA_ACTC924_SOLICTC','15:00','23:59','08:59:00'],
['CIP_CTC_SEND_VARREDURA_ACTC921_SOLICTC','19:00','20:00','01:00:00'],
['CIP_CTC_SEND_VARREDURA_DECPRZ','19:00','19:50','00:50:00'],
['CIP_CTC_SEND_VARREDURA_ACTC926_II','21:00','23:59','02:59:00'],
['CIP_CTC_GERA_ARQUIVO_TARIFACAO','23:59','04:59','05:00:00'],
['CIP_CTC_SEND_VARREDURA_ACTC901','18:30','04:59','10:29:00'],
['CIP_CTC_SEND_VARREDURA_ACTC921','00:00','04:59','04:59:00'],
['CIP_CTC_SEND_VARREDURA_ACTC922','19:00','04:59','09:59:00'],
['CIP_SCG_BAIXA_DECURSO_PRAZO_AGENDA','00:01','00:50','00:49:00'],
['CIP_SCG_BAIXA_ANTECIPACAO_AGENDA','05:00','06:00','01:00:00'],
['CIP_SCG_SEND_VARREDURA_ASCG004_AGENDA','15:00','15:30','00:30:00'],
['CIP_SCG_SEND_VARREDURA_ASCG008_AGENDA','15:00','15:30','00:30:00'],
['CIP_SCG_SEND_VARREDURA_ASCG020_AGENDA','15:00','15:30','00:30:00'],
['CIP_SCG_SEND_VARREDURA_ASCG002_AGENDA','19:00','23:29','04:29:00'],
['CIP_SCG_SEND_VARREDURA_ASCG009_AGENDA','19:00','23:29','04:29:00'],
['CIP_SCG_ENVIA_CD_SCG_AGENDA_FATURAMENTO_MENSAL','20:00','23:29','03:29:00'],
['CIP_SCG_EXPORT_SCG_AGENDA_FATURAMENTO_MENSAL','20:00','23:29','03:29:00'],
['CIP_SCG_GERAR_DATA_REFERENCIA_AGENDA','23:30','00:30','01:00:00'],
['CIP_SCG_UPDATE_GRADE_EVENTUAL','23:30','05:50','06:20:00'],
['CIP_SEC_NFP_ALTERAR_DATA_REFERENCIA','04:00','05:59','01:59:00'],
['CIP_SEC_NFP_GERAR_ARQUIVO_TARIFACAO','04:00','05:59','01:59:00'],
['CIP_STD_TARIFACAO','04:00','04:59','00:59:00'],
['CIP_STD_ALTERAR_DATA_REFERENCIA','04:00','04:59','00:59:00'],
['CIP_CQL_GERAR_DATA_REFERENCIA','00:00','03:59','03:59:00'],
['CIP_SEND_VARREDURA_ACQL001_22h00m','22:00','03:59','05:59:00'],
['CIP_SLC_VARREDURA_V_ASLC510_ENVIAR','04:30','05:29','00:59:00'],
['CIP_SLC_VARREDURA_V_ASLC520_CICLO01','05:05','05:29','00:24:00'],
['CIP_SLC_VARREDURA_V_ASLC510_DEV','10:10','10:59','00:49:00'],
['CIP_SLC_VARREDURA_V_ASLC510_ENVIAR_DEV','10:11','10:59','01:11:00'],
['CIP_SLC_VARREDURA_V_ASLC520_CICLO02','10:10','10:59','00:49:00'],
['CIP_SLC_VARREDURA_V_ASLC520_DEV','10:10','10:59','01:11:00'],
['CIP_SLC_VARREDURA_V_DEPRAZO','20:00','21:00','01:00:00'],
['CIP_SLC_VARREDURA_V_ASLC510','20:00','21:15','01:11:00'],
['CIP_SLC_GERA_DATA_REFERENCIA','23:30','23:35','00:05:00'],
['CIP_SLC_FATURAMENTO','23:59','01:00','01:01:00'],
['CIP_SLC_ENVIA_CIP_FATURAMENTO_MENSAL','01:00','01:10','01:11:00'],
)

############################################
# INSERINDO INFORMACAO NAS COLUNAS e LINHAS
############################################
jobs_format = workbook.add_format({
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
})

hrini_format = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
})

meta_format = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'font_color': 'red'
})

etime_format = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 12,
    'font_color': 'red'
})


hoje = datetime.datetime.today().day
hojev = hoje + 3
row = 2
col = 2
colvazio = 4
coldado = 5
ciclico = 0

# Iterate over the data and write it out row by row.
for job, hrini, meta, etime in (producao):
    worksheet.write(row, col, job, jobs_format)
    worksheet.write(row, col + 1, hrini, hrini_format)
    worksheet.write(row, col + 2, meta, meta_format)
    worksheet.write(row, col + 3, etime, etime_format)
    row += 1
###################################
#       FECHANDO A PLANILHA
###################################
workbook.close()
