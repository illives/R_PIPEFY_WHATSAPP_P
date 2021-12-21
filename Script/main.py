from Resources.resources import MessageModel
import os

def main():
    c = MessageModel()
    c.credencias()
    c.listar_cards()
    c.insert_new_cards()
    c.update_atributos()
    c.relatorio_geral()
    c.relatorio_novas_solicitacoes()
    c.relatorio_aprovados()
    c.relatorio_reprovados()
    c.novas_solicitacoes()
    c.aprovadas()
    c.reprovados()
    c.relatorio_diario()
    

if __name__ == '__main__':
    main()