from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()


# ---------- Bombeiro (utilizador do sistema) ----------
class Bombeiro(db.Model, UserMixin):
    __tablename__ = 'bombeiros'
    id = db.Column(db.Integer, primary_key=True)
    numero_interno = db.Column(db.String(20), unique=True, nullable=False)
    mecanografico = db.Column(db.String(20), unique=True, nullable=False)
    nome = db.Column(db.String(100), nullable=False)
    nomecompleto = db.Column(db.String(200))
    email = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    posto = db.Column(db.String(50), default='Bombeiro')
    tipo_bombeiro = db.Column(db.String(20), default='Voluntário')
    tipo_user = db.Column(db.String(20), default='User')
    telemovel = db.Column(db.String(20))
    resp_departamento = db.Column(db.String(50))
    ativo = db.Column(db.Boolean, default=True)

    # Relações com back_populates (todas bidirecionais)
    escalas = db.relationship('Escala', back_populates='bombeiro')
    trocas_origem = db.relationship('TrocaServico', foreign_keys='TrocaServico.bombeiro_origem_id', back_populates='bombeiro_origem')
    trocas_destino = db.relationship('TrocaServico', foreign_keys='TrocaServico.bombeiro_destino_id', back_populates='bombeiro_destino')
    dispensas = db.relationship('Dispensa', back_populates='bombeiro', cascade='all, delete-orphan')
    creditos = db.relationship('CreditoDispensa', back_populates='bombeiro')
    disponibilidades = db.relationship('Disponibilidade', back_populates='bombeiro')
    fardamentos = db.relationship('Fardamento', back_populates='bombeiro')
    fardamentos_atribuidos = db.relationship('FardamentoAtribuido', back_populates='bombeiro')
    ecins = db.relationship('Ecin', back_populates='bombeiro')
    deslocacoes = db.relationship('Deslocacao', back_populates='bombeiro')
    ferias = db.relationship('Ferias', foreign_keys='Ferias.bombeiro_id', back_populates='bombeiro')
    ferias_aprovadas = db.relationship('Ferias', foreign_keys='Ferias.aprovado_por', back_populates='aprovador')
    notas_criadas = db.relationship('Nota', back_populates='criador')
    mensagens_enviadas = db.relationship('MensagemCorreio', foreign_keys='MensagemCorreio.remetente_id', back_populates='remetente')
    mensagens_recebidas = db.relationship('MensagemCorreio', foreign_keys='MensagemCorreio.destinatario_id', back_populates='destinatario')
    checklists = db.relationship('Checklist', back_populates='bombeiro')
    checklists_feitos = db.relationship('ChecklistAmbulancia', back_populates='bombeiro')
    reunioes_criadas = db.relationship('Reuniao', back_populates='criador')
    notas_comando_criadas = db.relationship('NotaComando', back_populates='criador')
    avarias_reportadas = db.relationship('Avaria', back_populates='reportador')

    def get_id(self):
        return str(self.id)


# ---------- Viatura ----------
class Viatura(db.Model):
    __tablename__ = 'viaturas'
    id = db.Column(db.Integer, primary_key=True)
    matricula = db.Column(db.String(10), unique=True, nullable=False)
    tipo = db.Column(db.String(50))
    nomenclatura = db.Column(db.String(100))
    marca = db.Column(db.String(50))
    modelo = db.Column(db.String(50))
    ano = db.Column(db.Integer)
    estado = db.Column(db.String(20), default='operacional')

    avarias = db.relationship('Avaria', back_populates='viatura')
    checklists = db.relationship('Checklist', back_populates='viatura')
    oficina_registos = db.relationship('Oficina', back_populates='viatura')
    gestao_frota = db.relationship('GestaoFrota', back_populates='viatura', uselist=False)
    reposicoes_stock = db.relationship('StockAmbulancia', back_populates='ambulancia')
    checklists_ambulancia = db.relationship('ChecklistAmbulancia', back_populates='viatura')
    deslocacoes = db.relationship('Deslocacao', back_populates='viatura')


# ---------- Avaria ----------
class Avaria(db.Model):
    __tablename__ = 'avarias'
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(10), unique=True, nullable=False)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    descricao = db.Column(db.Text, nullable=False)
    data_reporte = db.Column(db.DateTime, default=datetime.utcnow)
    data_reparacao = db.Column(db.DateTime, nullable=True)
    reportado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    estado = db.Column(db.String(20), default='Pendente')
    kms = db.Column(db.Integer, nullable=True)
    responsavel_oficina = db.Column(db.Boolean, default=False)
    comando_verificado = db.Column(db.Boolean, default=False)
    parecer_resp = db.Column(db.Text, nullable=True)  # parecer do responsável da oficina
    urg_despacho = db.Column(db.String(30), nullable=True)  # Manutenção/Urgente/Não Urgente
    data_resp = db.Column(db.Date, nullable=True)  # data do parecer
    decisao_cmd = db.Column(db.Text, nullable=True)  # decisão do comando
    data_cmd = db.Column(db.Date, nullable=True)  # data da decisão

    viatura = db.relationship('Viatura', back_populates='avarias')
    reportador = db.relationship('Bombeiro', back_populates='avarias_reportadas')
    oficina_registos = db.relationship('Oficina', back_populates='avaria')


# ---------- Oficina ----------
class Oficina(db.Model):
    __tablename__ = 'oficina'
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(10), unique=True, nullable=False)
    data_registo = db.Column(db.DateTime, default=datetime.utcnow)
    nome_oficina = db.Column(db.String(100), nullable=False)
    data_recepcao = db.Column(db.Date, nullable=False)
    motivo = db.Column(db.Text)
    avaria_id = db.Column(db.Integer, db.ForeignKey('avarias.id'), nullable=True)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    kms = db.Column(db.Integer)
    inoperacional = db.Column(db.Boolean, default=False)
    descricao_oficina = db.Column(db.Text)
    n_orc_fat = db.Column(db.String(50))
    data_entrega = db.Column(db.Date)
    chefe_oficina = db.Column(db.Boolean, default=False)
    comando = db.Column(db.Boolean, default=False)
    operacional = db.Column(db.Boolean, default=False)
    estado = db.Column(db.String(20), default='Oficina')

    avaria = db.relationship('Avaria', back_populates='oficina_registos')
    viatura = db.relationship('Viatura', back_populates='oficina_registos')





# ---------- Gestão Frota ----------
class GestaoFrota(db.Model):
    __tablename__ = 'gestao_frota'
    id = db.Column(db.Integer, primary_key=True)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False, unique=True)
    inspecao_periodica = db.Column(db.Date, nullable=True)
    kms_ultima_revisao = db.Column(db.Integer)
    kms_proxima_revisao = db.Column(db.Integer)
    kms_pneus_dianteiros = db.Column(db.Integer)
    kms_pneus_trazeiros = db.Column(db.Integer)
    kms_correia = db.Column(db.Integer)
    outros_apontamentos = db.Column(db.Text)

    viatura = db.relationship('Viatura', back_populates='gestao_frota')


# ---------- Escala ----------
class Escala(db.Model):
    __tablename__ = 'escalas'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_inicio = db.Column(db.DateTime, nullable=False)
    data_fim = db.Column(db.DateTime, nullable=False)
    turno = db.Column(db.String(30), nullable=False)
    categoria = db.Column(db.String(50))
    funcao = db.Column(db.String(100))
    observacao = db.Column(db.String(100))

    bombeiro = db.relationship('Bombeiro', back_populates='escalas')


# ---------- Férias ----------
class Ferias(db.Model):
    __tablename__ = 'ferias'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_inicio = db.Column(db.Date, nullable=False)
    data_fim = db.Column(db.Date, nullable=False)
    estado = db.Column(db.String(20), default='Pendente')
    aprovado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    data_pedido = db.Column(db.DateTime, default=datetime.utcnow)

    bombeiro = db.relationship('Bombeiro', foreign_keys=[bombeiro_id], back_populates='ferias')
    aprovador = db.relationship('Bombeiro', foreign_keys=[aprovado_por], back_populates='ferias_aprovadas')


# ---------- Troca de Serviço ----------
class TrocaServico(db.Model):
    __tablename__ = 'trocas_servico'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_origem_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    bombeiro_destino_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data_origem = db.Column(db.Date, nullable=False)
    data_destino = db.Column(db.Date, nullable=False)
    turno_origem = db.Column(db.String(30))
    turno_destino = db.Column(db.String(30))
    motivo = db.Column(db.String(200))
    estado = db.Column(db.String(30), default='pendente_colega')
    data_pedido = db.Column(db.DateTime, default=datetime.utcnow)

    bombeiro_origem = db.relationship('Bombeiro', foreign_keys=[bombeiro_origem_id], back_populates='trocas_origem')
    bombeiro_destino = db.relationship('Bombeiro', foreign_keys=[bombeiro_destino_id], back_populates='trocas_destino')


# ---------- Dispensa ----------
class Dispensa(db.Model):
    __tablename__ = 'dispensas'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data_inicio = db.Column(db.Date, nullable=False)
    data_fim = db.Column(db.Date, nullable=False)
    motivo = db.Column(db.String(200))
    aprovada = db.Column(db.Boolean, default=False)
    categoria = db.Column(db.String(50))
    turno = db.Column(db.String(30))

    bombeiro = db.relationship('Bombeiro', back_populates='dispensas')
    creditos = db.relationship('CreditoDispensa', back_populates='dispensa')


# ---------- Créditos de Dispensa ----------
class CreditoDispensa(db.Model):
    __tablename__ = 'creditos_dispensa'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data = db.Column(db.Date, nullable=False)
    descricao = db.Column(db.String(200))
    horas = db.Column(db.Integer, default=8)
    observacao = db.Column(db.String(50), default='Não Gozado')
    dispensa_id = db.Column(db.Integer, db.ForeignKey('dispensas.id'), nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='creditos')
    dispensa = db.relationship('Dispensa', back_populates='creditos')


# ---------- Checklist de Viatura ----------
class Checklist(db.Model):
    __tablename__ = 'checklists'
    id = db.Column(db.Integer, primary_key=True)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.DateTime, default=datetime.utcnow)
    itens_verificados = db.Column(db.Text)
    observacoes = db.Column(db.Text)

    viatura = db.relationship('Viatura', back_populates='checklists')
    bombeiro = db.relationship('Bombeiro', back_populates='checklists')


# ---------- Fardamento ----------
class Fardamento(db.Model):
    __tablename__ = 'fardamentos'
    id = db.Column(db.Integer, primary_key=True)
    data_registo = db.Column(db.DateTime, default=datetime.utcnow)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    tipo = db.Column(db.String(50))
    nome = db.Column(db.String(100))
    descricao = db.Column(db.Text)
    tamanho = db.Column(db.String(20))
    motivo = db.Column(db.String(20))
    descricao_motivo = db.Column(db.Text)
    stock_id = db.Column(db.Integer, db.ForeignKey('stock_fardamento.id'), nullable=True)   # FK para o produto genérico
    estado = db.Column(db.String(20), default='Pedido')
    responsavel = db.Column(db.Boolean, default=False)
    comando = db.Column(db.Boolean, default=False)
    entregue = db.Column(db.Boolean, default=False)
    data_entrega = db.Column(db.DateTime, nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='fardamentos')
    # Relação unidirecional com StockFardamento (apenas leitura, sem back_populates)
    produto_stock = db.relationship('StockFardamento', foreign_keys=[stock_id])
    atribuicao = db.relationship('FardamentoAtribuido', back_populates='pedido', uselist=False,
                                 cascade='all, delete-orphan')


# ---------- Stock Fardamento ----------
class StockFardamento(db.Model):
    __tablename__ = 'stock_fardamento'
    id = db.Column(db.Integer, primary_key=True)
    codigo_farda = db.Column(db.String(20), unique=True, nullable=False)
    tipo = db.Column(db.String(50), default='Outro')
    nome = db.Column(db.String(100), nullable=False)
    descricao = db.Column(db.Text)

    # Relação com os itens no armazém (tamanhos/stocks)
    items_armazem = db.relationship('StockFardamentoArmazem', back_populates='produto', cascade='all, delete-orphan')


class StockFardamentoArmazem(db.Model):
    __tablename__ = 'stock_fardamento_armazem'
    __table_args__ = {'extend_existing': True}   # ← permite redefinir a tabela se já existir
    id = db.Column(db.Integer, primary_key=True)
    codigo_farda = db.Column(db.String(20), db.ForeignKey('stock_fardamento.codigo_farda'), nullable=False)
    sub_codigo_farda = db.Column(db.String(20), unique=True, nullable=False)
    tipo = db.Column(db.String(50), default='Outro')
    nome = db.Column(db.String(100), nullable=False)
    descricao = db.Column(db.Text)
    tamanho = db.Column(db.String(20))
    stock = db.Column(db.Integer, default=0)

    # Relação com o produto principal
    produto = db.relationship('StockFardamento', back_populates='items_armazem')


# ---------- Fardamento Atribuído ----------
class FardamentoAtribuido(db.Model):
    __tablename__ = 'fardamento_atribuido'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    tipo = db.Column(db.String(50))
    nome = db.Column(db.String(100))
    tamanho = db.Column(db.String(20))
    data_entrega = db.Column(db.Date, nullable=False)
    data_devolucao = db.Column(db.Date, nullable=True)
    estado = db.Column(db.String(20), default='Entregue')
    idpedido = db.Column(db.Integer, db.ForeignKey('fardamentos.id'), nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='fardamentos_atribuidos')
    pedido = db.relationship('Fardamento', back_populates='atribuicao')


# ---------- Tipo de Farda/Material ----------
class TipoFardaMaterial(db.Model):
    __tablename__ = 'tipos_farda_material'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), unique=True, nullable=False)
    categoria = db.Column(db.String(20), nullable=False, default='Farda')


# ---------- Disponibilidade ----------
class Disponibilidade(db.Model):
    __tablename__ = 'disponibilidades'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    turno_extra = db.Column(db.String(20))
    categoria = db.Column(db.String(30), nullable=True)
    confirmada = db.Column(db.Boolean, default=False)

    bombeiro = db.relationship('Bombeiro', back_populates='disponibilidades')


# ---------- ECIN / ELAC ----------
class Ecin(db.Model):
    __tablename__ = 'ecins'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    turno = db.Column(db.String(20), nullable=False)
    estado = db.Column(db.String(20), default='Pendente')
    funcao = db.Column(db.String(20), nullable=True)
    categoria = db.Column(db.String(30), nullable=True)
    valor = db.Column(db.Float, nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='ecins')
    mobilidades = db.relationship('Mobilidade', back_populates='ecin_original', cascade='all, delete-orphan')


class FarmaciaCentral(db.Model):
    __tablename__ = 'farmacia_central'
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20), unique=True, nullable=False)  # mesmo código do StockFarmacia
    categoria = db.Column(db.String(100), nullable=False)
    nome = db.Column(db.String(200), nullable=False)
    tamanho = db.Column(db.String(50))
    stock = db.Column(db.Integer, default=0)
    stock_minimo = db.Column(db.Integer, default=5)
    data_validade = db.Column(db.Date, nullable=True)
    ultima_atualizacao = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

# ---------- Stock Farmácia ----------
class StockFarmacia(db.Model):
    __tablename__ = 'stock_farmacia'
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(20), unique=True, nullable=False, default='SF0000')
    categoria = db.Column(db.String(100), nullable=False)
    nome = db.Column(db.String(200), nullable=False)
    tamanho = db.Column(db.String(50))
    stock = db.Column(db.Integer, default=0)
    infstock = db.Column(db.Integer, default=0)
    data_validade = db.Column(db.Date, nullable=True)
    data_atualizacao = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    saidas = db.relationship('StockAmbulancia', back_populates='produto', cascade='all, delete-orphan')
    checklist_itens = db.relationship('ChecklistAmbulanciaItem', back_populates='produto')


# ---------- Categoria Farmácia ----------
class CategoriaFarmacia(db.Model):
    __tablename__ = 'categorias_farmacia'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), unique=True, nullable=False)
    checklist = db.Column(db.Boolean, default=False)


# ---------- Stock Ambulância ----------
class StockAmbulancia(db.Model):
    __tablename__ = 'stock_ambulancia'
    id = db.Column(db.Integer, primary_key=True)
    ambulancia_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'))
    produto_id = db.Column(db.Integer, db.ForeignKey('stock_farmacia.id'))
    quantidade = db.Column(db.Integer, default=0)
    solicitante_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    responsavel_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    checklist_id = db.Column(db.Integer, db.ForeignKey('checklist_ambulancia.id'), nullable=True)
    confirmado = db.Column(db.Boolean, default=False)
    data = db.Column(db.DateTime, default=datetime.utcnow)

    ambulancia = db.relationship('Viatura', back_populates='reposicoes_stock')
    produto = db.relationship('StockFarmacia', back_populates='saidas')
    solicitante = db.relationship('Bombeiro', foreign_keys=[solicitante_id])
    responsavel = db.relationship('Bombeiro', foreign_keys=[responsavel_id])
    checklist = db.relationship('ChecklistAmbulancia', back_populates='reposicoes')


# ---------- Checklist Ambulância ----------
class ChecklistAmbulancia(db.Model):
    __tablename__ = 'checklist_ambulancia'
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.DateTime, default=datetime.utcnow)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    finalizado = db.Column(db.Boolean, default=False)

    viatura = db.relationship('Viatura', back_populates='checklists_ambulancia')
    bombeiro = db.relationship('Bombeiro', back_populates='checklists_feitos')
    itens = db.relationship('ChecklistAmbulanciaItem', back_populates='checklist', cascade='all, delete-orphan')
    reposicoes = db.relationship('StockAmbulancia', back_populates='checklist')


# ---------- Itens do Checklist Ambulância ----------
class ChecklistAmbulanciaItem(db.Model):
    __tablename__ = 'checklist_ambulancia_itens'
    id = db.Column(db.Integer, primary_key=True)
    checklist_id = db.Column(db.Integer, db.ForeignKey('checklist_ambulancia.id'), nullable=False)
    produto_id = db.Column(db.Integer, db.ForeignKey('stock_farmacia.id'), nullable=False)
    quantidade = db.Column(db.Integer, default=0)

    checklist = db.relationship('ChecklistAmbulancia', back_populates='itens')
    produto = db.relationship('StockFarmacia', back_populates='checklist_itens')


# ---------- Notas da Central ----------
class Nota(db.Model):
    __tablename__ = 'notas'
    id = db.Column(db.Integer, primary_key=True)
    criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    descricao = db.Column(db.Text, nullable=False)
    data_evento = db.Column(db.Date, nullable=True)

    criador = db.relationship('Bombeiro', back_populates='notas_criadas')


# ---------- Mensagens do Correio ----------
class MensagemCorreio(db.Model):
    __tablename__ = 'mensagens_correio'
    id = db.Column(db.Integer, primary_key=True)
    remetente_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    destinatario_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    departamento = db.Column(db.String(50), nullable=True)
    assunto = db.Column(db.String(150), nullable=False, default='Sem assunto')
    corpo = db.Column(db.Text, nullable=False)
    data_envio = db.Column(db.DateTime, default=datetime.utcnow)
    lida = db.Column(db.Boolean, default=False)
    apagada_remetente = db.Column(db.Boolean, default=False)
    apagada_destinatario = db.Column(db.Boolean, default=False)

    remetente = db.relationship('Bombeiro', foreign_keys=[remetente_id], back_populates='mensagens_enviadas')
    destinatario = db.relationship('Bombeiro', foreign_keys=[destinatario_id], back_populates='mensagens_recebidas')


# ---------- Reuniões ----------
class Reuniao(db.Model):
    __tablename__ = 'reunioes'
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.Date, nullable=False)
    hora = db.Column(db.String(10), nullable=True)
    assunto = db.Column(db.String(200), nullable=False)
    descricao = db.Column(db.Text, nullable=True)
    criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)

    criador = db.relationship('Bombeiro', back_populates='reunioes_criadas')


# ---------- Notas do Comando ----------
class NotaComando(db.Model):
    __tablename__ = 'notas_comando'
    id = db.Column(db.Integer, primary_key=True)
    criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    descricao = db.Column(db.Text, nullable=False)
    data_evento = db.Column(db.Date, nullable=True)

    criador = db.relationship('Bombeiro', back_populates='notas_comando_criadas')


# ---------- Deslocações ----------
class Deslocacao(db.Model):
    __tablename__ = 'deslocacoes'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    hora_inicio = db.Column(db.String(10), nullable=False)
    data_fim = db.Column(db.Date, nullable=True)          # NOVO
    hora_fim = db.Column(db.String(10), nullable=True)    # NOVO
    servico = db.Column(db.String(50), nullable=False)
    local_origem = db.Column(db.String(200))
    local_destino = db.Column(db.String(200))
    valor = db.Column(db.Float, nullable=True)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    n_servico = db.Column(db.String(20), nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='deslocacoes')
    viatura = db.relationship('Viatura', back_populates='deslocacoes')



class Mobilidade(db.Model):
    __tablename__ = 'mobilidades'

    id = db.Column(db.Integer, primary_key=True)
    ecin_original_id = db.Column(db.Integer, db.ForeignKey('ecins.id', ondelete='CASCADE'), nullable=False)
    bombeiro_substituto_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    horas = db.Column(db.Numeric(5,2), nullable=False)
    valor_pago = db.Column(db.Numeric(10,2), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)

    # Relações – ambas com back_populates
    ecin_original = db.relationship('Ecin', foreign_keys=[ecin_original_id], back_populates='mobilidades')
    bombeiro_substituto = db.relationship('Bombeiro', foreign_keys=[bombeiro_substituto_id])

class Monitor(db.Model):
    __tablename__ = 'monitor'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False, unique=True)
    fogos = db.Column(db.Boolean, default=True)
    google_maps = db.Column(db.Boolean, default=True)
    bombeiros_pt = db.Column(db.Boolean, default=True)
    ipma = db.Column(db.Boolean, default=True)
    pontoagua = db.Column(db.Boolean, default=True)  # ← nome correto: pontoagua

    bombeiro = db.relationship('Bombeiro', backref='monitor_config', uselist=False)


class PontoAgua(db.Model):
    __tablename__ = 'pontos_agua'  # atenção: nome da tabela
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    tipo = db.Column(db.String(50), default='Hidrante')
    latitude = db.Column(db.Float, nullable=False)
    longitude = db.Column(db.Float, nullable=False)
    freguesia = db.Column(db.String(100))  # ← campo correto
    descricao = db.Column(db.Text)
    capacidade = db.Column(db.String(50))
    criado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)

    criador = db.relationship('Bombeiro', foreign_keys=[criado_por])


# Modelo para guardar configurações do quadro operacional
class QuadroOperacional(db.Model):
    __tablename__ = 'quadro_operacional'
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.Date, nullable=False, unique=True)
    viatura_ecin_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    viatura_eip_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    viatura_inem_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    viatura_reserva_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    viatura_comando_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)

    # NOVOS CAMPOS
    motorista_inem_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    motorista_inem_numero = db.Column(db.String(20), nullable=True)
    motorista_inem_mec = db.Column(db.String(20), nullable=True)

    reserva_1_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    reserva_1_numero = db.Column(db.String(20), nullable=True)
    reserva_1_mec = db.Column(db.String(20), nullable=True)

    reserva_2_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)
    reserva_2_numero = db.Column(db.String(20), nullable=True)
    reserva_2_mec = db.Column(db.String(20), nullable=True)

    criado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)

    # Relações
    viatura_ecin = db.relationship('Viatura', foreign_keys=[viatura_ecin_id])
    viatura_eip = db.relationship('Viatura', foreign_keys=[viatura_eip_id])
    viatura_inem = db.relationship('Viatura', foreign_keys=[viatura_inem_id])
    viatura_reserva = db.relationship('Viatura', foreign_keys=[viatura_reserva_id])
    viatura_comando = db.relationship('Viatura', foreign_keys=[viatura_comando_id])

    motorista_inem = db.relationship('Bombeiro', foreign_keys=[motorista_inem_id])
    reserva_1 = db.relationship('Bombeiro', foreign_keys=[reserva_1_id])
    reserva_2 = db.relationship('Bombeiro', foreign_keys=[reserva_2_id])


    criador = db.relationship('Bombeiro', foreign_keys=[criado_por]).relationship('Bombeiro', foreign_keys=[criado_por])