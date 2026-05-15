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

    # Relações (todas com back_populates)
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
    ferias = db.relationship('Ferias', back_populates='bombeiro')
    notas_criadas = db.relationship('Nota', back_populates='criador')
    mensagens_enviadas = db.relationship('MensagemCorreio', foreign_keys='MensagemCorreio.remetente_id', back_populates='remetente')
    mensagens_recebidas = db.relationship('MensagemCorreio', foreign_keys='MensagemCorreio.destinatario_id', back_populates='destinatario')
    checklists = db.relationship('Checklist', back_populates='bombeiro')
    reunioes_criadas = db.relationship('Reuniao', back_populates='criador')
    notas_comando = db.relationship('NotaComando', back_populates='criador')

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

    avarias = db.relationship('Avaria', backref='viatura', lazy=True)
    checklists = db.relationship('Checklist', backref='viatura', lazy=True)


# ---------- Avaria de viatura ----------
class Avaria(db.Model):
    __tablename__ = 'avarias'
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(10), unique=True, nullable=False)  # novo campo
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    descricao = db.Column(db.Text, nullable=False)
    data_reporte = db.Column(db.DateTime, default=datetime.utcnow)
    data_reparacao = db.Column(db.DateTime, nullable=True)
    reportado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    estado = db.Column(db.String(20), default='Pendente')
    kms = db.Column(db.Integer, nullable=True)
    responsavel_oficina = db.Column(db.Boolean, default=False)
    comando_verificado = db.Column(db.Boolean, default=False)

    reportador = db.relationship('Bombeiro', backref='avarias_reportadas')


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
    estado = db.Column(db.String(20), default='Oficina')  # 'Oficina' ou 'Resolvido'

    avaria = db.relationship('Avaria', backref='oficina_registos')
    viatura = db.relationship('Viatura', backref='oficina_registos')

# ---------- GestaoFrota ----------
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

    viatura = db.relationship('Viatura', backref=db.backref('gestao_frota', uselist=False))


# ---------- Escala de serviço ----------
class Escala(db.Model):
    __tablename__ = 'escalas'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_inicio = db.Column(db.Date, nullable=False)      # ← antes era DateTime
    data_fim = db.Column(db.Date, nullable=False)          # ← idem
    turno = db.Column(db.String(20))
    categoria = db.Column(db.String(30), default='Bombeiro')
    funcao = db.Column(db.String(50))
    observacao = db.Column(db.String(100), nullable=True)

    bombeiro = db.relationship('Bombeiro', backref='escalas')

class Ferias(db.Model):
    __tablename__ = 'ferias'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_inicio = db.Column(db.Date, nullable=False)
    data_fim = db.Column(db.Date, nullable=False)
    estado = db.Column(db.String(20), default='Pendente')  # Pendente, Aprovado, Rejeitado
    data_pedido = db.Column(db.DateTime, default=datetime.utcnow)
    aprovado_por = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)

    bombeiro = db.relationship('Bombeiro', foreign_keys=[bombeiro_id], backref='ferias_pedidas')
    aprovador = db.relationship('Bombeiro', foreign_keys=[aprovado_por])




# ---------- Troca de serviço ----------
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

    # Relações
    bombeiro_origem = db.relationship('Bombeiro', foreign_keys=[bombeiro_origem_id], back_populates='trocas_origem')
    bombeiro_destino = db.relationship('Bombeiro', foreign_keys=[bombeiro_destino_id], back_populates='trocas_destino')


# ---------- Dispensa de serviço ----------
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

    # Relações
    bombeiro = db.relationship('Bombeiro', back_populates='dispensas')
    creditos = db.relationship('CreditoDispensa', back_populates='dispensa')

class CreditoDispensa(db.Model):
    __tablename__ = 'creditos_dispensa'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'))
    data = db.Column(db.Date, nullable=False)
    descricao = db.Column(db.String(200))
    horas = db.Column(db.Integer, default=8)
    observacao = db.Column(db.String(50), default='Não Gozado')
    dispensa_id = db.Column(db.Integer, db.ForeignKey('dispensas.id'), nullable=True)

    # Relações
    bombeiro = db.relationship('Bombeiro', back_populates='creditos')
    dispensa = db.relationship('Dispensa', back_populates='creditos')


# ---------- Checklist ----------
class Checklist(db.Model):
    __tablename__ = 'checklists'
    id = db.Column(db.Integer, primary_key=True)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.DateTime, default=datetime.utcnow)
    itens_verificados = db.Column(db.Text)
    observacoes = db.Column(db.Text)

    bombeiro = db.relationship('Bombeiro', backref='checklists')



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
    stock_id = db.Column(db.Integer, db.ForeignKey('stock_fardamento.id'), nullable=True)
    estado = db.Column(db.String(20), default='Pedido')
    responsavel = db.Column(db.Boolean, default=False)
    comando = db.Column(db.Boolean, default=False)
    entregue = db.Column(db.Boolean, default=False)
    data_entrega = db.Column(db.DateTime, nullable=True)

    bombeiro = db.relationship('Bombeiro', back_populates='fardamentos')
    stock = db.relationship('StockFardamento', back_populates='fardamentos')
    atribuicao = db.relationship('FardamentoAtribuido', back_populates='pedido', uselist=False)

# ---------- Stock Fardamento ----------
class StockFardamento(db.Model):
    __tablename__ = 'stock_fardamento'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    descricao = db.Column(db.Text)
    tamanho = db.Column(db.String(20))
    tipo = db.Column(db.String(50), default='Outro')
    stock = db.Column(db.Integer, default=0)

    fardamentos = db.relationship('Fardamento', back_populates='stock', lazy=True)



# ---------- Disponibilidade ----------
class Disponibilidade(db.Model):
    __tablename__ = 'disponibilidades'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    turno_extra = db.Column(db.String(20))
    categoria = db.Column(db.String(30), nullable=True)   # novo campo
    confirmada = db.Column(db.Boolean, default=False)

    bombeiro = db.relationship('Bombeiro', back_populates='disponibilidades')

# ---------- Ecin ----------

class Ecin(db.Model):
    __tablename__ = 'ecins'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    turno = db.Column(db.String(20), nullable=False)
    estado = db.Column(db.String(20), default='Pendente')   # agora guarda a legenda após escalamento
    funcao = db.Column(db.String(20), nullable=True)        # Motorista, Chefe, Guarnição
    categoria = db.Column(db.String(30), nullable=True)     # ECIN, ELAC
    valor = db.Column(db.Float, nullable=True)  # novo campo

    bombeiro = db.relationship('Bombeiro', backref='ecins')

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

    # Relação com StockAmbulancia (sem backref problemático)
    saidas = db.relationship('StockAmbulancia', back_populates='produto', cascade='all, delete-orphan')



class CategoriaFarmacia(db.Model):
    __tablename__ = 'categorias_farmacia'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), unique=True, nullable=False)
    checklist = db.Column(db.Boolean, default=False)   # novo campo



    # ---------- Stock Ambulancia ----------
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

    ambulancia = db.relationship('Viatura', backref='reposicoes_stock')
    produto = db.relationship('StockFarmacia', back_populates='saidas')
    solicitante = db.relationship('Bombeiro', foreign_keys=[solicitante_id])
    responsavel = db.relationship('Bombeiro', foreign_keys=[responsavel_id])
    checklist = db.relationship('ChecklistAmbulancia')  # ← sem backref

# ---------- Checklist Ambulancia ----------

class ChecklistAmbulancia(db.Model):
    __tablename__ = 'checklist_ambulancia'
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.DateTime, default=datetime.utcnow)
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=False)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)   # quem fez a checklist
    finalizado = db.Column(db.Boolean, default=False)

    viatura = db.relationship('Viatura', backref='checklists_ambulancia')
    bombeiro = db.relationship('Bombeiro', backref='checklists_feitos')
    itens = db.relationship('ChecklistAmbulanciaItem', back_populates='checklist', lazy=True, cascade='all, delete-orphan')
    reposicoes = db.relationship('StockAmbulancia', back_populates='checklist')


class ChecklistAmbulanciaItem(db.Model):
    __tablename__ = 'checklist_ambulancia_itens'
    id = db.Column(db.Integer, primary_key=True)
    checklist_id = db.Column(db.Integer, db.ForeignKey('checklist_ambulancia.id'), nullable=False)
    produto_id = db.Column(db.Integer, db.ForeignKey('stock_farmacia.id'), nullable=False)
    quantidade = db.Column(db.Integer, default=0)   # preenchido depois

    checklist = db.relationship('ChecklistAmbulancia', back_populates='itens')
    produto = db.relationship('StockFarmacia', backref='checklist_itens')

class Nota(db.Model):
        __tablename__ = 'notas'
        id = db.Column(db.Integer, primary_key=True)
        criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
        data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
        descricao = db.Column(db.Text, nullable=False)
        data_evento = db.Column(db.Date, nullable=True)  # data do evento, opcional

        criador = db.relationship('Bombeiro', backref='notas_criadas')


class MensagemCorreio(db.Model):
    __tablename__ = 'mensagens_correio'
    id = db.Column(db.Integer, primary_key=True)
    remetente_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    destinatario_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=True)   # se for para um bombeiro
    departamento = db.Column(db.String(50), nullable=True)   # se for para um departamento inteiro
    assunto = db.Column(db.String(150), nullable=False, default='Sem assunto')
    corpo = db.Column(db.Text, nullable=False)
    data_envio = db.Column(db.DateTime, default=datetime.utcnow)
    lida = db.Column(db.Boolean, default=False)
    apagada_remetente = db.Column(db.Boolean, default=False)   # para o remetente poder apagar da vista
    apagada_destinatario = db.Column(db.Boolean, default=False) # para o destinatário poder apagar

    remetente = db.relationship('Bombeiro', foreign_keys=[remetente_id], backref='mensagens_enviadas')
    destinatario = db.relationship('Bombeiro', foreign_keys=[destinatario_id], backref='mensagens_recebidas')
    # Sem a linha de relationship – o acesso a bombeiro será feito via backref da classe Bombeiro


#-------------------Fardamento--------------
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
    idpedido = db.Column(db.Integer, db.ForeignKey('fardamentos.id'), nullable=True)   # NOVO CAMPO

    bombeiro = db.relationship('Bombeiro', backref='fardamentos_atribuidos')
    pedido = db.relationship('Fardamento', backref='atribuicao')   # relação com o pedido original
    pedido = db.relationship('Fardamento', back_populates='atribuicao')


class TipoFardaMaterial(db.Model):
    __tablename__ = 'tipos_farda_material'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), unique=True, nullable=False)
    categoria = db.Column(db.String(20), nullable=False, default='Farda')  # 'Farda' ou 'Material'


class Reuniao(db.Model):
    __tablename__ = 'reunioes'
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.Date, nullable=False)
    hora = db.Column(db.String(10), nullable=True)          # Formato HH:MM, opcional
    assunto = db.Column(db.String(200), nullable=False)
    descricao = db.Column(db.Text, nullable=True)
    criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)

    criador = db.relationship('Bombeiro', backref='reunioes_criadas')

class NotaComando(db.Model):
    __tablename__ = 'notas_comando'
    id = db.Column(db.Integer, primary_key=True)
    criador_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    descricao = db.Column(db.Text, nullable=False)
    data_evento = db.Column(db.Date, nullable=True)

    criador = db.relationship('Bombeiro', backref='notas_comando_criadas')


class Deslocacao(db.Model):
    __tablename__ = 'deslocacoes'
    id = db.Column(db.Integer, primary_key=True)
    bombeiro_id = db.Column(db.Integer, db.ForeignKey('bombeiros.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    hora_inicio = db.Column(db.String(10), nullable=False)  # HH:MM
    servico = db.Column(db.String(50), nullable=False)  # "C. Doentes", "Evacuação", "Retorno", "Urgência"
    local_origem = db.Column(db.String(200))
    local_destino = db.Column(db.String(200))
    valor = db.Column(db.Float, nullable=True)  # visível apenas a Admin/Comando/Secretaria
    viatura_id = db.Column(db.Integer, db.ForeignKey('viaturas.id'), nullable=True)
    n_servico = db.Column(db.String(20), nullable=True)

    bombeiro = db.relationship('Bombeiro', backref='deslocacoes')
    viatura = db.relationship('Viatura', backref='deslocacoes')