import os
from datetime import datetime, timedelta, date
from io import BytesIO
import calendar
from flask import Flask, render_template, redirect, url_for, request, flash, send_file, jsonify
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill
from sqlalchemy import func

from models import db, Bombeiro, Viatura, Avaria, Escala, TrocaServico, Dispensa, Checklist, Fardamento, Disponibilidade, CreditoDispensa, Oficina, GestaoFrota, StockFardamento, Ecin, StockFarmacia, StockAmbulancia, ChecklistAmbulancia, CategoriaFarmacia, ChecklistAmbulanciaItem, Nota, MensagemCorreio, FardamentoAtribuido, Reuniao, NotaComando

app = Flask(__name__)

# Chave secreta (obrigatória para sessões e formulários)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'chave-local-insegura')

# Desativar o rastreio de modificações (poupa memória)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Base de dados: usa PostgreSQL se a variável DATABASE_URL existir (Render),
# senão mantém o SQLite local para desenvolvimento.
basedir = os.path.abspath(os.path.dirname(__file__))
db_url = os.environ.get('DATABASE_URL', 'sqlite:///' + os.path.join(basedir, 'bombeiros.db'))
if db_url:
    if db_url.startswith('postgres://'):
        db_url = db_url.replace('postgres://', 'postgresql+psycopg://', 1)
    elif db_url.startswith('postgresql://'):
        db_url = db_url.replace('postgresql://', 'postgresql+psycopg://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_url

# Inicializar a extensão com a aplicação
db.init_app(app)


login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Bombeiro, int(user_id))


# ---------- Criação da BD e admin inicial ----------
@app.before_request
def create_tables():
    db.create_all()
    if not Bombeiro.query.filter_by(email='admin@quartel.pt').first():
        admin = Bombeiro(
            numero_interno='B000',
            mecanografico='M000',
            nome='Administrador',
            email='admin@quartel.pt',
            password_hash=generate_password_hash('admin123'),
            posto='Comandante',
            ativo=True,
            tipo_user='Admin',
            tipo_bombeiro='Voluntário'
        )
        db.session.add(admin)
        if not Viatura.query.first():
            v1 = Viatura(matricula='12-AB-34', tipo='ABSC', nomenclatura='ABSC 01', marca='Mercedes', modelo='Atego', ano=2020)
            v2 = Viatura(matricula='56-CD-78', tipo='VLCI', nomenclatura='VLCI 02', marca='MAN', modelo='TGM', ano=2019)
            db.session.add_all([v1, v2])
        db.session.commit()


# ---------- Autenticação ----------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        mecanografico = request.form['mecanografico'].strip()
        password = request.form['password']
        user = Bombeiro.query.filter_by(mecanografico=mecanografico).first()
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            flash('Sessão iniciada com sucesso.', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('dashboard'))
        else:
            flash('Mecanográfico ou password inválidos.', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Sessão terminada.', 'info')
    return redirect(url_for('login'))


def turno_para_horas(turno_str):
    """Converte string do turno (ex: '1 - 00h/08h') em (hora_inicio, hora_fim)."""
    try:
        partes = turno_str.split('-')[1].strip().split('/')
        inicio = partes[0].replace('h', ':00')
        fim = partes[1].replace('h', ':00')
        return inicio, fim
    except Exception:
        return '08:00', '20:00'


from sqlalchemy import func   # Certifique-se de que está no topo do app.py

def gerar_html_dia(data, categoria_filtro=''):
    # Escalas do dia
    query = Escala.query.filter(
        func.date(Escala.data_inicio) <= data,
        func.date(Escala.data_fim) >= data
    )
    if categoria_filtro:
        query = query.filter(Escala.categoria == categoria_filtro)
    escalas = query.order_by(Escala.categoria, Escala.turno, Escala.data_inicio).all()

    # IDs dos bombeiros que aparecem nas escalas (filtrados ou não)
    ids_bombeiros_escalados = {e.bombeiro_id for e in escalas}

    # Agrupar por categoria
    escalas_por_categoria = {}
    for esc in escalas:
        cat = esc.categoria or 'Outros'
        if cat not in escalas_por_categoria:
            escalas_por_categoria[cat] = []
        escalas_por_categoria[cat].append(esc)

    cores_categorias = {
        'Motorista': '#fd7e14',
        'Socorrista': '#20c997',
        'Centralista': '#0d6efd',
        'EIP': '#6f42c1',
        'ECIN': '#dc3545',
        'ELAC': '#fd7e14',
        'Piquete': '#198754',
        'Bombeiro': '#0dcaf0'
    }

    # Dispensas do dia
    dispensas_do_dia = Dispensa.query.filter(
        Dispensa.data_inicio <= data,
        Dispensa.data_fim >= data,
        Dispensa.aprovada == True
    )
    if categoria_filtro:
        dispensas_do_dia = dispensas_do_dia.filter(Dispensa.bombeiro_id.in_(ids_bombeiros_escalados))
    dispensas_do_dia = dispensas_do_dia.order_by(Dispensa.data_inicio).all()
    bombeiros_com_dispensa = {d.bombeiro_id for d in dispensas_do_dia}

    # Trocas do dia
    trocas_do_dia = TrocaServico.query.filter(
        (TrocaServico.data_origem == data) | (TrocaServico.data_destino == data),
        TrocaServico.estado == 'aprovada'
    )
    if categoria_filtro:
        trocas_do_dia = trocas_do_dia.filter(
            (TrocaServico.bombeiro_origem_id.in_(ids_bombeiros_escalados)) |
            (TrocaServico.bombeiro_destino_id.in_(ids_bombeiros_escalados))
        )
    trocas_do_dia = trocas_do_dia.order_by(TrocaServico.data_origem).all()

    html = ''
    if escalas_por_categoria:
        html += '<h6 class="text-secondary mt-2">Escalas</h6>'
        for cat, lista in escalas_por_categoria.items():
            cor = cores_categorias.get(cat, '#6c757d')
            html += f'<h6 style="color: {cor};"><i class="bi bi-person-badge me-1"></i>{cat}</h6>'
            html += '<ul class="list-group list-group-flush mb-2">'
            for e in lista:
                estilo_linha = ''
                if e.bombeiro_id in bombeiros_com_dispensa:
                    estilo_linha = 'style="background-color: #ffe5cc;"'
                html += f'<li class="list-group-item py-1 d-flex justify-content-between" {estilo_linha}>'
                html += f'<span>{e.bombeiro.nome} ({e.bombeiro.mecanografico})</span>'
                html += f'<span>{e.turno}</span>'
                html += '</li>'
            html += '</ul>'
    else:
        html += '<p class="text-muted">Nenhuma escala para este dia nesta categoria.</p>'

    if trocas_do_dia:
        html += '<hr><h6 class="text-secondary">Trocas</h6><ul class="list-group list-group-flush mb-2">'
        for t in trocas_do_dia:
            if t.data_origem == data:
                texto = f'{t.bombeiro_origem.nome} estava escalado para este dia, mas trocou com {t.bombeiro_destino.nome}.'
            else:
                texto = f'{t.bombeiro_destino.nome} estava escalado para este dia, mas trocou com {t.bombeiro_origem.nome}.'
            html += f'<li class="list-group-item py-1">{texto}</li>'
        html += '</ul>'
    else:
        html += '<hr><p class="text-muted">Sem trocas neste dia.</p>'

    if dispensas_do_dia:
        html += '<hr><h6 class="text-secondary">Dispensas</h6><ul class="list-group list-group-flush mb-2">'
        for d in dispensas_do_dia:
            html += f'<li class="list-group-item py-1">{d.bombeiro.nome} ({d.data_inicio.strftime("%d/%m/%Y")} a {d.data_fim.strftime("%d/%m/%Y")})</li>'
        html += '</ul>'
    else:
        html += '<hr><p class="text-muted">Sem dispensas neste dia.</p>'

    return html


# ---------- Dashboard ----------
@app.route('/')
@login_required
def dashboard():
    return render_template('dashboard.html')


# ---------- Gestão de Bombeiros ----------
@app.route('/bombeiros', methods=['GET', 'POST'])
@login_required
def gerir_bombeiros():
    if current_user.tipo_user != 'Admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        numero = request.form['numero_interno']
        mecanografico = request.form['mecanografico']
        nome = request.form['nome']
        nomecompleto = request.form.get('nomecompleto', '')
        email = request.form['email']
        password = request.form['password']
        posto = request.form['posto']
        tipo_bombeiro = request.form.get('tipo_bombeiro', 'Voluntário')
        tipo_user = request.form['tipo_user']
        telemovel = request.form.get('telemovel', '')
        resp_departamento = request.form.get('resp_departamento', '')

        if Bombeiro.query.filter((Bombeiro.email == email) |
                                 (Bombeiro.numero_interno == numero) |
                                 (Bombeiro.mecanografico == mecanografico)).first():
            flash('Email, número interno ou mecanográfico já existe.', 'warning')
            return redirect(url_for('gerir_bombeiros'))

        if telemovel and Bombeiro.query.filter_by(telemovel=telemovel).first():
            flash('Telemóvel já está associado a outro bombeiro.', 'warning')
            return redirect(url_for('gerir_bombeiros'))

        novo = Bombeiro(
            numero_interno=numero,
            mecanografico=mecanografico,
            nome=nome,
            nomecompleto=nomecompleto if nomecompleto else None,
            email=email,
            password_hash=generate_password_hash(password),
            posto=posto,
            tipo_bombeiro=tipo_bombeiro,
            tipo_user=tipo_user,
            telemovel=telemovel if telemovel else None,
            resp_departamento=resp_departamento if resp_departamento else None,
            ativo=True
        )
        db.session.add(novo)
        db.session.commit()
        flash('Bombeiro criado com sucesso!', 'success')
        return redirect(url_for('gerir_bombeiros'))

    # GET – listagem com filtro e ordenação
    pesquisa = request.args.get('pesquisa', '').strip()
    query = Bombeiro.query
    if pesquisa:
        query = query.filter(
            (Bombeiro.nome.ilike(f'%{pesquisa}%')) |
            (Bombeiro.mecanografico.ilike(f'%{pesquisa}%'))
        )
    bombeiros = query.order_by(Bombeiro.nome.asc()).all()
    return render_template('bombeiros.html', bombeiros=bombeiros, pesquisa=pesquisa)

@app.route('/bombeiros/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_bombeiro(id):
    if current_user.tipo_user != 'Admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))

    bombeiro = Bombeiro.query.get_or_404(id)

    if request.method == 'POST':
        novo_email = request.form['email']
        novo_numero = request.form['numero_interno']
        novo_mec = request.form['mecanografico']
        novo_telemovel = request.form.get('telemovel', '')

        conflito = Bombeiro.query.filter(
            (Bombeiro.id != id) &
            ((Bombeiro.email == novo_email) |
             (Bombeiro.numero_interno == novo_numero) |
             (Bombeiro.mecanografico == novo_mec))
        ).first()
        if conflito:
            flash('Email, nº interno ou mecanográfico já em uso.', 'warning')
            return redirect(url_for('gerir_bombeiros'))

        if novo_telemovel:
            telemovel_conflito = Bombeiro.query.filter(
                Bombeiro.id != id, Bombeiro.telemovel == novo_telemovel
            ).first()
            if telemovel_conflito:
                flash('Telemóvel já associado a outro bombeiro.', 'warning')
                return redirect(url_for('gerir_bombeiros'))

        bombeiro.numero_interno = novo_numero
        bombeiro.mecanografico = novo_mec
        bombeiro.nome = request.form['nome']
        bombeiro.nomecompleto = request.form.get('nomecompleto', '')
        bombeiro.email = novo_email
        bombeiro.telemovel = novo_telemovel if novo_telemovel else None
        bombeiro.posto = request.form['posto']
        bombeiro.tipo_bombeiro = request.form.get('tipo_bombeiro', 'Voluntário')
        bombeiro.tipo_user = request.form['tipo_user']
        bombeiro.resp_departamento = request.form.get('resp_departamento', '')
        bombeiro.ativo = request.form.get('ativo') == 'on'

        nova_password = request.form.get('password')
        if nova_password:
            bombeiro.password_hash = generate_password_hash(nova_password)

        db.session.commit()
        flash('Bombeiro atualizado com sucesso!', 'success')
        return redirect(url_for('gerir_bombeiros'))

    return redirect(url_for('gerir_bombeiros'))


@app.route('/bombeiros/apagar/<int:id>')
@login_required
def apagar_bombeiro(id):
    if current_user.tipo_user != 'Admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))

    bombeiro = Bombeiro.query.get_or_404(id)
    if bombeiro.id == current_user.id:
        flash('Não pode apagar o seu próprio utilizador.', 'warning')
        return redirect(url_for('gerir_bombeiros'))

    db.session.delete(bombeiro)
    db.session.commit()
    flash(f'Bombeiro {bombeiro.nome} removido.', 'info')
    return redirect(url_for('gerir_bombeiros'))

# ---------- Exportar Bombeiros ----------

@app.route('/bombeiros/exportar')
@login_required
def exportar_bombeiros():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    bombeiros = Bombeiro.query.order_by(Bombeiro.numero_interno).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bombeiros"

    cabecalhos = ['Nº Interno', 'Mecanográfico', 'Nome', 'Nome Completo', 'Email', 'Telemóvel',
                  'Posto', 'Tipo Bombeiro', 'Resp. Departamento', 'Tipo Utilizador', 'Ativo']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for b in bombeiros:
        ws.append([
            b.numero_interno,
            b.mecanografico,
            b.nome,
            b.nomecompleto or '',
            b.email,
            b.telemovel or '',
            b.posto,
            b.tipo_bombeiro,
            b.resp_departamento or '',
            b.tipo_user,
            'Sim' if b.ativo else 'Não'
        ])

    col_widths = [12, 16, 25, 35, 30, 15, 20, 16, 22, 16, 8]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='bombeiros.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- Importar Bombeiros ----------
@app.route('/bombeiros/importar', methods=['POST'])
@login_required
def importar_bombeiros():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('gerir_bombeiros'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('gerir_bombeiros'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido. Use .xlsx.', 'danger')
        return redirect(url_for('gerir_bombeiros'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('gerir_bombeiros'))

    linhas_importadas = 0
    erros = []

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue

        try:
            numero = str(row[0]).strip() if row[0] else None
            mecanografico = str(row[1]).strip() if len(row) > 1 and row[1] else None
            nome = str(row[2]).strip() if len(row) > 2 and row[2] else None
            nomecompleto = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            email = str(row[4]).strip().lower() if len(row) > 4 and row[4] else None
            telemovel = str(row[5]).strip() if len(row) > 5 and row[5] else None
            posto = str(row[6]).strip() if len(row) > 6 and row[6] else 'Bombeiro'
            tipo_bombeiro = str(row[7]).strip() if len(row) > 7 and row[7] else 'Voluntário'
            departamento = str(row[8]).strip() if len(row) > 8 and row[8] else None
            tipo_user = str(row[9]).strip() if len(row) > 9 and row[9] else 'User'
            ativo_str = str(row[10]).strip().lower() if len(row) > 10 and row[10] else 'sim'
            ativo = ativo_str == 'sim'
        except Exception:
            erros.append(f'Linha {row_num}: dados inválidos.')
            continue

        if not numero or not mecanografico or not nome or not email:
            erros.append(f'Linha {row_num}: campos obrigatórios em falta.')
            continue

        existente = Bombeiro.query.filter(
            (Bombeiro.numero_interno == numero) |
            (Bombeiro.mecanografico == mecanografico) |
            (Bombeiro.email == email)
        ).first()
        if existente:
            erros.append(f'Linha {row_num}: já existe um bombeiro com o Nº Interno, Mecanográfico ou Email indicado.')
            continue

        novo = Bombeiro(
            numero_interno=numero,
            mecanografico=mecanografico,
            nome=nome,
            nomecompleto=nomecompleto if nomecompleto else None,
            email=email,
            password_hash=generate_password_hash('123456'),
            telemovel=telemovel if telemovel else None,
            posto=posto,
            tipo_bombeiro=tipo_bombeiro,
            resp_departamento=departamento,
            tipo_user=tipo_user,
            ativo=ativo
        )
        db.session.add(novo)
        linhas_importadas += 1

    db.session.commit()

    if erros:
        flash(f'{linhas_importadas} importados. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} bombeiros importados com sucesso!', 'success')
    return redirect(url_for('gerir_bombeiros'))


# ---------- Viaturas ----------
@app.route('/viaturas')
@login_required
def listar_viaturas():
    viaturas = Viatura.query.order_by(Viatura.tipo, Viatura.matricula).all()
    # Contagens case‑insensitive
    total_operacionais = Viatura.query.filter(func.lower(Viatura.estado) == 'operacional').count()
    total_inoperacionais = Viatura.query.filter(func.lower(Viatura.estado) == 'inoperacional').count()
    total_manutencao = Viatura.query.filter(func.lower(Viatura.estado) == 'manutenção').count()
    total_geral = Viatura.query.count()
    return render_template('viaturas.html',
                           viaturas=viaturas,
                           total_operacionais=total_operacionais,
                           total_inoperacionais=total_inoperacionais,
                           total_manutencao=total_manutencao,
                           total_geral=total_geral)


@app.route('/viaturas/adicionar', methods=['POST'])
@login_required
def adicionar_viatura():
    if current_user.tipo_user != 'Admin':
        flash('Apenas administradores podem adicionar viaturas.', 'danger')
        return redirect(url_for('listar_viaturas'))

    matricula = request.form['matricula']
    tipo = request.form['tipo']
    nomenclatura = request.form['nomenclatura']
    marca = request.form['marca']
    modelo = request.form['modelo']
    ano = request.form['ano']

    nova = Viatura(matricula=matricula, tipo=tipo, nomenclatura=nomenclatura,
                   marca=marca, modelo=modelo, ano=ano)
    db.session.add(nova)
    db.session.commit()

    gestao = GestaoFrota(viatura_id=nova.id)
    db.session.add(gestao)
    db.session.commit()

    flash('Viatura adicionada.', 'success')
    return redirect(url_for('listar_viaturas'))


@app.route('/viaturas/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_viatura(id):
    if current_user.tipo_user != 'Admin':
        flash('Apenas administradores podem editar viaturas.', 'danger')
        return redirect(url_for('listar_viaturas'))

    viatura = Viatura.query.get_or_404(id)

    if request.method == 'POST':
        matricula = request.form['matricula']
        existente = Viatura.query.filter(Viatura.matricula == matricula, Viatura.id != id).first()
        if existente:
            flash('Matrícula já existe.', 'warning')
            return redirect(url_for('listar_viaturas'))

        viatura.matricula = matricula
        viatura.tipo = request.form['tipo']
        viatura.nomenclatura = request.form['nomenclatura']
        viatura.marca = request.form['marca']
        viatura.modelo = request.form['modelo']
        viatura.ano = request.form['ano']
        viatura.estado = request.form['estado']
        db.session.commit()
        flash('Viatura atualizada.', 'success')
        return redirect(url_for('listar_viaturas'))

    return redirect(url_for('listar_viaturas'))


@app.route('/viaturas/apagar/<int:id>')
@login_required
def apagar_viatura(id):
    if current_user.tipo_user != 'Admin':
        flash('Apenas administradores podem apagar viaturas.', 'danger')
        return redirect(url_for('listar_viaturas'))

    viatura = Viatura.query.get_or_404(id)
    db.session.delete(viatura)
    db.session.commit()
    flash(f'Viatura {viatura.matricula} removida.', 'info')
    return redirect(url_for('listar_viaturas'))


# ---------- Exportar Viaturas ----------
@app.route('/viaturas/exportar')
@login_required
def exportar_viaturas():
    viaturas = Viatura.query.order_by(Viatura.matricula).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Viaturas"

    cabecalhos = ['Matrícula', 'Tipo', 'Nomenclatura', 'Marca', 'Modelo', 'Ano', 'Estado']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for v in viaturas:
        ws.append([v.matricula, v.tipo, v.nomenclatura, v.marca, v.modelo, v.ano, v.estado])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='viaturas.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ---------- Importar Viaturas ----------

@app.route('/viaturas/importar', methods=['POST'])
@login_required
def importar_viaturas():
    if current_user.tipo_user != 'Admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('listar_viaturas'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('listar_viaturas'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('listar_viaturas'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('listar_viaturas'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            matricula = str(row[0]).strip() if row[0] else None
            tipo = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            nomenclatura = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            marca = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            modelo = str(row[4]).strip() if len(row) > 4 and row[4] else ''
            ano = int(row[5]) if len(row) > 5 and row[5] else None
            estado = str(row[6]).strip().lower() if len(row) > 6 and row[6] else 'operacional'
        except Exception:
            erros.append(f'Linha {row_num}: dados inválidos.')
            continue

        if not matricula:
            erros.append(f'Linha {row_num}: matrícula obrigatória.')
            continue
        if Viatura.query.filter_by(matricula=matricula).first():
            erros.append(f'Linha {row_num}: matrícula {matricula} já existe.')
            continue

        nova = Viatura(matricula=matricula, tipo=tipo, nomenclatura=nomenclatura,
                       marca=marca, modelo=modelo, ano=ano, estado=estado)
        db.session.add(nova)
        # Criar também registo na Gestão de Frota
        gestao = GestaoFrota(viatura_id=nova.id)
        db.session.add(gestao)
        linhas_importadas += 1

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importadas. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} viaturas importadas com sucesso.', 'success')
    return redirect(url_for('listar_viaturas'))


# ---------- Gestão Frota ----------
@app.route('/gestao-frota')
@login_required
def gestao_frota():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    # Criar registos de gestão em falta para qualquer viatura que ainda não tenha um
    viaturas_sem_gestao = Viatura.query.filter(~Viatura.gestao_frota.has()).all()
    for v in viaturas_sem_gestao:
        nova_gestao = GestaoFrota(viatura_id=v.id)
        db.session.add(nova_gestao)
    if viaturas_sem_gestao:
        db.session.commit()

    # Consulta que junta todas as viaturas com os seus dados de gestão (left join)
    registos = db.session.query(Viatura, GestaoFrota)\
                         .outerjoin(GestaoFrota, Viatura.id == GestaoFrota.viatura_id)\
                         .order_by(Viatura.matricula).all()

    viaturas = Viatura.query.order_by(Viatura.matricula).all()
    return render_template('gestao_frota.html', registos=registos, viaturas=viaturas)


# ---------- Editar Gestão Frota ----------
@app.route('/gestao-frota/editar/<int:id>', methods=['POST'])
@login_required
def editar_gestao_frota(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('gestao_frota'))

    registo = GestaoFrota.query.get_or_404(id)

    # Campos de data
    inspecao_str = request.form.get('inspecao_periodica')
    registo.inspecao_periodica = datetime.strptime(inspecao_str, '%Y-%m-%d').date() if inspecao_str else None

    # Campos numéricos (convertidos para int, None se vazios)
    registo.kms_ultima_revisao = request.form.get('kms_ultima_revisao', type=int)
    registo.kms_proxima_revisao = request.form.get('kms_proxima_revisao', type=int)
    registo.kms_pneus_dianteiros = request.form.get('kms_pneus_dianteiros', type=int)
    registo.kms_pneus_trazeiros = request.form.get('kms_pneus_trazeiros', type=int)
    registo.kms_correia = request.form.get('kms_correia', type=int)
    registo.outros_apontamentos = request.form.get('outros_apontamentos', '')

    db.session.commit()
    flash('Registo atualizado.', 'success')
    return redirect(url_for('gestao_frota'))

# ---------- Exportar Gestão Frota ----------

@app.route('/gestao-frota/exportar')
@login_required
def exportar_gestao_frota():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('gestao_frota'))

    registos = GestaoFrota.query.join(Viatura).order_by(Viatura.matricula).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gestao Frota"

    cabecalhos = ['ID', 'Matrícula', 'Nomenclatura', 'Marca/Modelo', 'Ano',
                  'Inspeção Periódica', 'Kms Última Revisão', 'Kms Próxima Revisão',
                  'Kms Pneus Diant.', 'Kms Pneus Tras.', 'Kms Correia', 'Outros Apontamentos']
    ws.append(cabecalhos)

    # Estilo cabeçalho
    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for r in registos:
        v = r.viatura
        ws.append([
            r.id,
            v.matricula,
            v.nomenclatura,
            f"{v.marca} {v.modelo}",
            v.ano,
            r.inspecao_periodica.strftime('%d/%m/%Y') if r.inspecao_periodica else '',
            r.kms_ultima_revisao or '',
            r.kms_proxima_revisao or '',
            r.kms_pneus_dianteiros or '',
            r.kms_pneus_trazeiros or '',
            r.kms_correia or '',
            r.outros_apontamentos or ''
        ])

    # Ajustar larguras
    col_widths = [5, 12, 18, 20, 6, 14, 14, 14, 12, 12, 12, 25]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='gestao_frota.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ---------- Importar Gestão Frota ----------

@app.route('/gestao-frota/importar', methods=['POST'])
@login_required
def importar_gestao_frota():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('gestao_frota'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('gestao_frota'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('gestao_frota'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('gestao_frota'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            matricula = str(row[1]).strip() if len(row) > 1 and row[1] else None
            if not matricula:
                erros.append(f'Linha {row_num}: matrícula em falta.')
                continue
            viatura = Viatura.query.filter_by(matricula=matricula).first()
            if not viatura:
                erros.append(f'Linha {row_num}: viatura {matricula} não encontrada.')
                continue

            gestao = GestaoFrota.query.filter_by(viatura_id=viatura.id).first()
            if not gestao:
                # Cria se não existir
                gestao = GestaoFrota(viatura_id=viatura.id)
                db.session.add(gestao)

            # Atualizar campos
            # Inspeção periódica (coluna 6, índice 5)
            inspecao_str = str(row[5]).strip() if len(row) > 5 and row[5] else None
            if inspecao_str:
                gestao.inspecao_periodica = _parse_data(inspecao_str)
            else:
                gestao.inspecao_periodica = None

            gestao.kms_ultima_revisao = int(row[6]) if len(row) > 6 and row[6] else None
            gestao.kms_proxima_revisao = int(row[7]) if len(row) > 7 and row[7] else None
            gestao.kms_pneus_dianteiros = int(row[8]) if len(row) > 8 and row[8] else None
            gestao.kms_pneus_trazeiros = int(row[9]) if len(row) > 9 and row[9] else None
            gestao.kms_correia = int(row[10]) if len(row) > 10 and row[10] else None
            gestao.outros_apontamentos = str(row[11]).strip() if len(row) > 11 and row[11] else None

            linhas_importadas += 1
        except Exception as e:
            erros.append(f'Linha {row_num}: erro {str(e)}')
            continue

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importadas com {len(erros)} erro(s).', 'warning')
    else:
        flash(f'{linhas_importadas} registos importados com sucesso.', 'success')
    return redirect(url_for('gestao_frota'))

# ---------- Imprimir Gestão Frota ----------
@app.route('/gestao-frota/imprimir/<int:viatura_id>')
@login_required
def imprimir_gestao_frota(viatura_id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('gestao_frota'))

    viatura = Viatura.query.get_or_404(viatura_id)
    gestao = GestaoFrota.query.filter_by(viatura_id=viatura_id).first()
    return render_template('imprimir_frota.html', viatura=viatura, gestao=gestao)


# ---------- Avarias ----------
@app.route('/avarias', methods=['GET', 'POST'])
@login_required
def avarias():
    tab = request.args.get('tab', 'registo')

    # Se for POST, é criação de uma nova avaria (mantido como antes)
    if request.method == 'POST':
        viatura_id = request.form['viatura_id']
        descricao = request.form['descricao']
        kms = request.form.get('kms', '')

        # Gerar código automático
        ultima_avaria = Avaria.query.order_by(Avaria.id.desc()).first()
        proximo = 1
        if ultima_avaria:
            try:
                proximo = int(ultima_avaria.codigo[2:]) + 1
            except Exception:
                pass
        codigo = f"AV{proximo:04d}"

        nova = Avaria(
            codigo=codigo,
            viatura_id=viatura_id,
            descricao=descricao,
            reportado_por=current_user.id,
            kms=int(kms) if kms else None,
            responsavel_oficina=False,
            comando_verificado=False
        )
        db.session.add(nova)
        db.session.commit()
        flash(f'Avaria {codigo} registada.', 'success')
        return redirect(url_for('avarias', tab='registo'))

    # ---- GET ----
    # Inicializar variáveis de filtro
    filtro_viatura_id = request.args.get('viatura_id', type=int)
    filtro_mes = request.args.get('mes', type=int)
    filtro_ano = request.args.get('ano', type=int)
    filtro_estado = request.args.get('estado', '')

    # Lista de todas as viaturas para os filtros
    todas_viaturas = Viatura.query.order_by(Viatura.matricula).all()

    # Dados para a aba Registo (sempre disponíveis para o formulário e tabela principal)
    viaturas = Viatura.query.all()  # para o modal de criação
    avarias_lista = Avaria.query.order_by(Avaria.data_reporte.desc()).limit(100).all()  # opcional: limite para não pesar

    # Dados para a aba Histórico (apenas quando a aba está ativa)
    historico_avarias = []
    if tab == 'historico':
        query = Avaria.query
        if filtro_viatura_id:
            query = query.filter_by(viatura_id=filtro_viatura_id)
        if filtro_mes:
            query = query.filter(db.extract('month', Avaria.data_reporte) == filtro_mes)
        if filtro_ano:
            query = query.filter(db.extract('year', Avaria.data_reporte) == filtro_ano)
        if filtro_estado:
            query = query.filter_by(estado=filtro_estado)
        historico_avarias = query.order_by(Avaria.data_reporte.desc()).all()

    return render_template('avarias.html',
                           avarias=avarias_lista,
                           viaturas=viaturas,
                           todas_viaturas=todas_viaturas,
                           historico_avarias=historico_avarias,
                           filtro_viatura_id=filtro_viatura_id,
                           filtro_mes=filtro_mes,
                           filtro_ano=filtro_ano,
                           filtro_estado=filtro_estado,
                           tab=tab)

@app.route('/avarias/atualizar/<int:id>', methods=['POST'])
@login_required
def atualizar_avaria(id):
    avaria = Avaria.query.get_or_404(id)

    permitido = (current_user.tipo_user == 'Admin' or
                 current_user.resp_departamento in ['Oficina', 'Comando'])
    if not permitido:
        flash('Sem permissão para alterar.', 'danger')
        return redirect(url_for('avarias'))

    novo_estado = request.form.get('estado')
    if novo_estado in ['Pendente', 'Analisar', 'Resolvido']:
        avaria.estado = novo_estado

    kms = request.form.get('kms', '')
    if kms:
        avaria.kms = int(kms)

    if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Oficina':
        avaria.responsavel_oficina = (request.form.get('resp_oficina') == 'on')

    if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Comando':
        avaria.comando_verificado = (request.form.get('comando_verificado') == 'on')

    db.session.commit()
    flash('Avaria atualizada.', 'success')
    return redirect(url_for('avarias'))


@app.route('/avarias/historico/<int:viatura_id>')
@login_required
def historico_avaria(viatura_id):
    viatura = Viatura.query.get_or_404(viatura_id)
    avarias_lista = Avaria.query.filter_by(viatura_id=viatura_id)\
                                .order_by(Avaria.data_reporte.desc()).all()
    return render_template('_historico_avarias.html', viatura=viatura, avarias=avarias_lista)


# ---------- Exportar Avarias ----------
@app.route('/avarias/exportar')
@login_required
def exportar_avarias():
    avarias_lista = Avaria.query.order_by(Avaria.data_reporte.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Avarias"

    cabecalhos = ['Código', 'Viatura', 'Descrição', 'Reportado por', 'Kms',
                  'Resp. Oficina', 'Comando', 'Estado', 'Data Reporte']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for a in avarias_lista:
        ws.append([
            a.codigo,
            a.viatura.matricula,
            a.descricao,
            a.reportador.nome,
            a.kms or '',
            'Sim' if a.responsavel_oficina else 'Não',
            'Sim' if a.comando_verificado else 'Não',
            a.estado,
            a.data_reporte.strftime('%d/%m/%Y %H:%M') if a.data_reporte else ''
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='avarias.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- Importar Avarias ----------

@app.route('/avarias/importar', methods=['POST'])
@login_required
def importar_avarias():
    if current_user.tipo_user != 'Admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('avarias'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('avarias'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('avarias'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('avarias'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            codigo = str(row[0]).strip() if row[0] else None
            viatura_matricula = str(row[1]).strip() if len(row) > 1 and row[1] else None
            descricao = str(row[2]).strip() if len(row) > 2 and row[2] else None
            reportador_mec = str(row[3]).strip() if len(row) > 3 and row[3] else None
            kms = int(row[4]) if len(row) > 4 and row[4] else None
            resp_oficina = str(row[5]).strip().lower() == 'sim' if len(row) > 5 and row[5] else False
            comando = str(row[6]).strip().lower() == 'sim' if len(row) > 6 and row[6] else False
            estado = str(row[7]).strip() if len(row) > 7 and row[7] else 'Pendente'
            data_reporte = _parse_data(row[8]) if len(row) > 8 and row[8] else datetime.utcnow()
        except Exception:
            erros.append(f'Linha {row_num}: dados inválidos.')
            continue

        if not descricao or not viatura_matricula:
            erros.append(f'Linha {row_num}: descrição e matrícula obrigatórias.')
            continue
        viatura = Viatura.query.filter_by(matricula=viatura_matricula).first()
        if not viatura:
            erros.append(f'Linha {row_num}: viatura {viatura_matricula} não encontrada.')
            continue
        bombeiro = Bombeiro.query.filter_by(mecanografico=reportador_mec).first() if reportador_mec else None
        if not bombeiro and reportador_mec:
            erros.append(f'Linha {row_num}: mecanográfico {reportador_mec} não encontrado.')
            continue

        # Gerar código único se necessário
        if not codigo:
            ultima = Avaria.query.order_by(Avaria.id.desc()).first()
            proximo = 1 if not ultima else int(ultima.codigo[2:]) + 1
            codigo = f"AV{proximo:04d}"

        nova = Avaria(
            codigo=codigo,
            viatura_id=viatura.id,
            descricao=descricao,
            reportado_por=bombeiro.id if bombeiro else current_user.id,
            kms=kms,
            responsavel_oficina=resp_oficina,
            comando_verificado=comando,
            estado=estado,
            data_reporte=data_reporte if data_reporte else datetime.utcnow()
        )
        db.session.add(nova)
        linhas_importadas += 1

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importadas. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} avarias importadas com sucesso.', 'success')
    return redirect(url_for('avarias'))

# ---------- Oficina ----------
@app.route('/oficina', methods=['GET', 'POST'])
@login_required
def oficina():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    tab = request.args.get('tab', 'registo')

    if request.method == 'POST':
        # Gerar código
        ultimo = Oficina.query.order_by(Oficina.id.desc()).first()
        proximo = 1
        if ultimo:
            try:
                proximo = int(ultimo.codigo[2:]) + 1
            except Exception:
                pass
        codigo = f"OF{proximo:04d}"

        avaria_id_str = request.form.get('avaria_id')
        avaria_id = int(avaria_id_str) if avaria_id_str else None

        viatura_id = None
        kms = None
        if avaria_id:
            avaria = Avaria.query.get(avaria_id)
            if avaria:
                viatura_id = avaria.viatura_id
                kms = avaria.kms
        if not viatura_id:
            viatura_id = request.form.get('viatura_id_manual', type=int)
        if not kms:
            kms = request.form.get('kms', type=int)

        nome_oficina = request.form['nome_oficina']
        data_recepcao = datetime.strptime(request.form['data_recepcao'], '%Y-%m-%d').date()
        motivo = request.form.get('motivo', '')
        descricao_oficina = request.form.get('descricao_oficina', '')
        n_orc_fat = request.form.get('n_orc_fat', '')
        data_entrega_str = request.form.get('data_entrega')
        data_entrega = datetime.strptime(data_entrega_str, '%Y-%m-%d').date() if data_entrega_str else None

        inoperacional = request.form.get('inoperacional') == 'on'
        chefe_oficina = request.form.get('chefe_oficina') == 'on'
        comando = request.form.get('comando') == 'on'
        operacional = request.form.get('operacional') == 'on'

        estado = 'Oficina'
        if operacional or (chefe_oficina and comando):
            estado = 'Resolvido'

        nova_oficina = Oficina(
            codigo=codigo,
            nome_oficina=nome_oficina,
            data_recepcao=data_recepcao,
            motivo=motivo,
            avaria_id=avaria_id,
            viatura_id=viatura_id,
            kms=kms,
            inoperacional=inoperacional,
            descricao_oficina=descricao_oficina,
            n_orc_fat=n_orc_fat,
            data_entrega=data_entrega,
            chefe_oficina=chefe_oficina,
            comando=comando,
            operacional=operacional,
            estado=estado
        )
        db.session.add(nova_oficina)

        # Atualizar estado da viatura
        viatura = Viatura.query.get(viatura_id)
        if viatura:
            if operacional:
                viatura.estado = 'operacional'
                nova_oficina.inoperacional = False
            elif inoperacional:
                viatura.estado = 'Inoperacional'

        db.session.commit()
        flash(f'Registo de oficina {codigo} criado.', 'success')
        return redirect(url_for('oficina', tab='registo'))

    # ---- GET ----
    filtro_viatura_id = request.args.get('viatura_id', type=int)
    filtro_nome_oficina = request.args.get('nome_oficina', '')
    filtro_mes = request.args.get('mes', type=int)
    filtro_ano = request.args.get('ano', type=int)
    filtro_estado = request.args.get('estado', '')

    todas_viaturas = Viatura.query.order_by(Viatura.matricula).all()
    nomes_oficina = [row[0] for row in db.session.query(Oficina.nome_oficina).distinct().all()]

    # Listas para a aba Registo
    avarias_analisar = Avaria.query.filter_by(estado='Analisar').all()
    viaturas = Viatura.query.all()
    registos = Oficina.query.order_by(Oficina.id.desc()).all()

    # Histórico filtrado
    historico_oficina = []
    if tab == 'historico':
        query = Oficina.query
        if filtro_viatura_id:
            query = query.filter_by(viatura_id=filtro_viatura_id)
        if filtro_nome_oficina:
            query = query.filter_by(nome_oficina=filtro_nome_oficina)
        if filtro_mes:
            query = query.filter(db.extract('month', Oficina.data_recepcao) == filtro_mes)
        if filtro_ano:
            query = query.filter(db.extract('year', Oficina.data_recepcao) == filtro_ano)
        if filtro_estado:
            query = query.filter_by(estado=filtro_estado)
        historico_oficina = query.order_by(Oficina.data_registo.desc()).all()

    return render_template('oficina.html',
                           registos=registos,
                           avarias_analisar=avarias_analisar,
                           viaturas=viaturas,
                           todas_viaturas=todas_viaturas,
                           nomes_oficina=nomes_oficina,
                           historico_oficina=historico_oficina,
                           filtro_viatura_id=filtro_viatura_id,
                           filtro_nome_oficina=filtro_nome_oficina,
                           filtro_mes=filtro_mes,
                           filtro_ano=filtro_ano,
                           filtro_estado=filtro_estado,
                           tab=tab)

# ---------- Histórico Oficina ----------
@app.route('/oficina/historico/<int:viatura_id>')
@login_required
def historico_oficina(viatura_id):
    viatura = Viatura.query.get_or_404(viatura_id)
    registos = Oficina.query.filter_by(viatura_id=viatura_id)\
                            .order_by(Oficina.data_registo.desc()).all()
    return render_template('_historico_oficina.html', viatura=viatura, registos=registos)

# ---------- Apagar Reg. Oficina ----------
@app.route('/oficina/apagar/<int:id>')
@login_required
def apagar_registo_oficina(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('oficina'))

    registo = Oficina.query.get_or_404(id)
    db.session.delete(registo)
    db.session.commit()
    flash(f'Registo {registo.codigo} removido.', 'info')
    return redirect(url_for('oficina'))

# ---------- Editar Reg. Oficina ----------

@app.route('/oficina/editar/<int:id>', methods=['POST'])
@login_required
def editar_registo_oficina(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('oficina'))

    registo = Oficina.query.get_or_404(id)

    # Atualizar campos básicos
    registo.nome_oficina = request.form['nome_oficina']
    registo.data_recepcao = datetime.strptime(request.form['data_recepcao'], '%Y-%m-%d').date()
    registo.motivo = request.form.get('motivo', '')
    registo.descricao_oficina = request.form.get('descricao_oficina', '')
    registo.n_orc_fat = request.form.get('n_orc_fat', '')
    data_entrega_str = request.form.get('data_entrega')
    registo.data_entrega = datetime.strptime(data_entrega_str, '%Y-%m-%d').date() if data_entrega_str else None

    # Avaria e viatura/kms
    avaria_id_str = request.form.get('avaria_id')
    if avaria_id_str:
        avaria_id = int(avaria_id_str)
        avaria = Avaria.query.get(avaria_id)
        if avaria:
            registo.avaria_id = avaria_id
            registo.viatura_id = avaria.viatura_id
            registo.kms = avaria.kms
    else:
        registo.avaria_id = None
        # Se não veio da avaria, permite escolher viatura diretamente (caso exista no formulário)
        viatura_id = request.form.get('viatura_id', type=int)
        if viatura_id:
            registo.viatura_id = viatura_id
        kms = request.form.get('kms', type=int)
        if kms is not None:
            registo.kms = kms

    # Checkboxes
    inoperacional = request.form.get('inoperacional') == 'on'
    chefe_oficina = request.form.get('chefe_oficina') == 'on'
    comando = request.form.get('comando') == 'on'
    operacional = request.form.get('operacional') == 'on'

    registo.inoperacional = inoperacional
    registo.chefe_oficina = chefe_oficina
    registo.comando = comando
    registo.operacional = operacional

    # Determinar estado
    if operacional or (chefe_oficina and comando):
        registo.estado = 'Resolvido'
        registo.operacional = True   # força se ambos chefe+comando foram marcados
    else:
        registo.estado = 'Oficina'

    # Atualizar estado da viatura
    viatura = Viatura.query.get(registo.viatura_id)
    if viatura:
        if inoperacional:
            viatura.estado = 'Inoperacional'
        elif operacional or registo.estado == 'Resolvido':
            viatura.estado = 'operacional'
            registo.inoperacional = False  # retirar visto inoperacional no registo
        # Se não é inoperacional nem operacional, mantém o estado atual

    db.session.commit()
    flash(f'Registo {registo.codigo} atualizado.', 'success')
    return redirect(url_for('oficina'))


# ---------- Exportar Oficina ----------
@app.route('/oficina/exportar')
@login_required
def exportar_oficina():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('oficina'))

    registos = Oficina.query.order_by(Oficina.data_registo.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Oficina"

    cabecalhos = ['Código', 'Data Registo', 'Nome Oficina', 'Data Recepção', 'Motivo',
                  'Nº Avaria', 'Viatura', 'Kms', 'Inoperacional', 'Descrição Oficina',
                  'Nº Orç/Fat', 'Data Entrega', 'Chefe Oficina', 'Comando', 'Operacional', 'Estado']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for r in registos:
        ws.append([
            r.codigo,
            r.data_registo.strftime('%d/%m/%Y %H:%M') if r.data_registo else '',
            r.nome_oficina,
            r.data_recepcao.strftime('%d/%m/%Y') if r.data_recepcao else '',
            r.motivo or '',
            r.avaria.codigo if r.avaria else '',
            r.viatura.matricula if r.viatura else '',
            r.kms or '',
            'Sim' if r.inoperacional else 'Não',
            r.descricao_oficina or '',
            r.n_orc_fat or '',
            r.data_entrega.strftime('%d/%m/%Y') if r.data_entrega else '',
            'Sim' if r.chefe_oficina else 'Não',
            'Sim' if r.comando else 'Não',
            'Sim' if r.operacional else 'Não',
            r.estado
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='oficina.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- Importar Oficina ----------
@app.route('/oficina/importar', methods=['POST'])
@login_required
def importar_oficina():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Oficina']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('oficina'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('oficina'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '':
        flash('Ficheiro vazio.', 'warning')
        return redirect(url_for('oficina'))
    if not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('oficina'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('oficina'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            nome_oficina = str(row[0]).strip() if row[0] else None
            data_recepcao = _parse_data(row[1]) if len(row) > 1 else None
            motivo = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            avaria_codigo = str(row[3]).strip() if len(row) > 3 and row[3] else None
            viatura_matricula = str(row[4]).strip() if len(row) > 4 and row[4] else None
            kms = int(row[5]) if len(row) > 5 and row[5] else 0
            inoperacional = str(row[6]).strip().lower() == 'sim' if len(row) > 6 and row[6] else False
            descricao = str(row[7]).strip() if len(row) > 7 and row[7] else ''
            n_orc_fat = str(row[8]).strip() if len(row) > 8 and row[8] else ''
            data_entrega = _parse_data(row[9]) if len(row) > 9 and row[9] else None
            chefe_oficina = str(row[10]).strip().lower() == 'sim' if len(row) > 10 and row[10] else False
            comando = str(row[11]).strip().lower() == 'sim' if len(row) > 11 and row[11] else False
            operacional = str(row[12]).strip().lower() == 'sim' if len(row) > 12 and row[12] else False
        except Exception as e:
            erros.append(f'Linha {row_num}: erro ao interpretar dados.')
            continue

        if not nome_oficina or not data_recepcao:
            erros.append(f'Linha {row_num}: nome oficina e data recepção obrigatórios.')
            continue

        viatura = Viatura.query.filter_by(matricula=viatura_matricula).first() if viatura_matricula else None
        avaria = Avaria.query.filter_by(codigo=avaria_codigo).first() if avaria_codigo else None

        if not viatura:
            erros.append(f'Linha {row_num}: viatura não encontrada.')
            continue

        # Gerar código único
        ultimo = Oficina.query.order_by(Oficina.id.desc()).first()
        proximo = 1 if not ultimo else int(ultimo.codigo[2:]) + 1
        codigo = f"OF{proximo:04d}"

        nova = Oficina(
            codigo=codigo,
            nome_oficina=nome_oficina,
            data_recepcao=data_recepcao,
            motivo=motivo,
            avaria_id=avaria.id if avaria else None,
            viatura_id=viatura.id,
            kms=kms,
            inoperacional=inoperacional,
            descricao_oficina=descricao,
            n_orc_fat=n_orc_fat,
            data_entrega=data_entrega,
            chefe_oficina=chefe_oficina,
            comando=comando,
            operacional=operacional,
            estado='Resolvido' if (chefe_oficina and comando) else 'Oficina'
        )
        db.session.add(nova)
        linhas_importadas += 1

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importadas com {len(erros)} erro(s).', 'warning')
    else:
        flash(f'{linhas_importadas} importadas com sucesso.', 'success')
    return redirect(url_for('oficina'))


# ---------- Escala ----------
@app.route('/escala')
@login_required
def escala():
    mes = request.args.get('mes', type=int, default=date.today().month)
    categoria = request.args.get('categoria', '')
    mecanografico = request.args.get('mecanografico', '')
    turno_filtro = request.args.get('turno', '')
    dia_filtro = request.args.get('dia', type=int)  # dia do mês

    query = Escala.query.join(Bombeiro)

    if mes:
        query = query.filter(db.extract('month', Escala.data_inicio) == mes)
    if categoria:
        query = query.filter(Escala.categoria == categoria)
    if mecanografico:
        query = query.filter(Bombeiro.mecanografico == mecanografico)
    if turno_filtro:
        query = query.filter(Escala.turno == turno_filtro)
    if dia_filtro and mes:
        hoje = date.today()
        data_ref = date(hoje.year, mes, dia_filtro)
        query = query.filter(
            func.date(Escala.data_inicio) <= data_ref,
            func.date(Escala.data_fim) >= data_ref
        )

    if current_user.tipo_user != 'Admin':
        query = query.filter(Escala.bombeiro_id == current_user.id)

    # Obter todas as escalas que correspondem aos filtros
    escalas_brutas = query.order_by(Escala.data_inicio.asc()).all()

    # ---------- ORDENAÇÃO PERSONALIZADA ----------
    categorias_ordem = ['Motorista', 'Socorrista', 'Centralista', 'EIP',
                        'ECIN', 'ELAC', 'Piquete']

    prioridades = {
        'Motorista': ['Luís Matias','Jorge Pereira','José Soldado','José Seco','Pedro Fernandes',
              'David Charrinho','Fábio Leirinha','Soeiro Mendes','Ana Marzia',
              'Filipe Martins','Eric Nobre'],
        'Socorrista': ['José Rodrigues','Paulo Branquinho','Sabrina Fernandes'],
        'Centralista': ['Mariana Charrinho','Ruben Ramos','António Pequito'],
        'EIP': ['José Fernandes','João Mateus','Tiago Bizarro','João Carita','João Silva']
    }
    for cat in prioridades:
        prioridades[cat] = {nome: i for i, nome in enumerate(prioridades[cat])}

    def chave_ordenacao(esc):
        # 1º critério: data (convertendo para date)
        data_ini = esc.data_inicio.date() if hasattr(esc.data_inicio, 'date') else esc.data_inicio
        # 2º critério: ordem da categoria
        cat = esc.categoria if esc.categoria else 'Outros'
        ordem_cat = categorias_ordem.index(cat) if cat in categorias_ordem else len(categorias_ordem)
        nome = esc.bombeiro.nome.strip()
        # 3º critério: posição do nome (se categoria tiver lista) ou turno+nome
        if cat in prioridades:
            pos_nome = prioridades[cat].get(nome, len(prioridades[cat]))
            return (data_ini, ordem_cat, pos_nome, nome)
        else:
            return (data_ini, ordem_cat, esc.turno, nome)

    escalas = sorted(escalas_brutas, key=chave_ordenacao)

    # Cartões de resumo (opcional, mantidos para compatibilidade com o template)
    total_escalas = len(escalas)
    total_bombeiros = len(set(e.bombeiro_id for e in escalas))
    categorias_presentes = set(e.categoria for e in escalas if e.categoria)
    total_categorias = len(categorias_presentes)
    total_turnos = len(set(e.turno for e in escalas))

    # Atividade do mês (para mini-calendário)
    dias_com_escalas = []
    if mes:
        ano_ref = datetime.now().year
        for dia in range(1, 32):
            try:
                data_ref = date(ano_ref, mes, dia)
            except ValueError:
                break
            tem = Escala.query.filter(
                db.extract('month', Escala.data_inicio) == mes,
                func.date(Escala.data_inicio) <= data_ref,
                func.date(Escala.data_fim) >= data_ref
            ).first() is not None
            if tem:
                dias_com_escalas.append(dia)

    # Trocas e dispensas de hoje (para indicadores)
    hoje = date.today()
    trocas_hoje = set()
    for t in TrocaServico.query.filter(
        (TrocaServico.data_origem == hoje) | (TrocaServico.data_destino == hoje),
        TrocaServico.estado == 'aprovada'
    ).all():
        trocas_hoje.add(t.bombeiro_origem_id)
        trocas_hoje.add(t.bombeiro_destino_id)

    dispensas_hoje = set(d.bombeiro_id for d in Dispensa.query.filter(
        Dispensa.data_inicio <= hoje,
        Dispensa.data_fim >= hoje,
        Dispensa.aprovada == True
    ).all())

    meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
    categorias = ['Motorista','Socorrista','Centralista','EIP','ECIN','ELAC','Piquete','Bombeiro']
    turnos = ['1 - 00h/08h','2 - 08h/16h','3 - 16h/24h','4 - 11h/19h','5 - 10h/18h','6 - 07h/19h','7 - 19h/07h','8 - 08h/20h','9 - 20h/08h']
    bombeiros_ativos = Bombeiro.query.filter_by(ativo=True).all()
    mecanograficos_ativos = [b.mecanografico for b in bombeiros_ativos]

    return render_template('escala.html',
                           escalas=escalas,
                           meses=meses,
                           categorias=categorias,
                           turnos=turnos,
                           mes_atual=mes,
                           categoria_atual=categoria,
                           turno_atual_filtro=turno_filtro,
                           dia_filtro=dia_filtro,
                           mecanografico_atual=mecanografico,
                           bombeiros_ativos=bombeiros_ativos,
                           mecanograficos_ativos=mecanograficos_ativos,
                           total_escalas=total_escalas,
                           total_bombeiros=total_bombeiros,
                           total_categorias=total_categorias,
                           total_turnos=total_turnos,
                           dias_com_escalas=dias_com_escalas,
                           trocas_hoje=trocas_hoje,
                           dispensas_hoje=dispensas_hoje,
                           hoje=hoje,
                           now=date.today())

@app.route('/escala/imprimir-mes')
@login_required
def imprimir_escala_mes():
    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)
    if not mes or not ano:
        hoje = date.today()
        mes = hoje.month
        ano = hoje.year

    # Feriados fixos em Portugal
    feriados = []
    feriados_fixos = {
        (1, 1): "Ano Novo",
        (4, 25): "Dia da Liberdade",
        (5, 1): "Dia do Trabalhador",
        (6, 10): "Dia de Portugal",
        (8, 15): "Assunção de Nossa Senhora",
        (10, 5): "Implantação da República",
        (11, 1): "Todos os Santos",
        (12, 1): "Restauração da Independência",
        (12, 8): "Imaculada Conceição",
        (12, 25): "Natal"
    }
    for (m, d), nome in feriados_fixos.items():
        try:
            feriados.append(date(ano, m, d))
        except ValueError:
            pass

    # Obter todas as escalas do mês, excluindo ECIN, ELAC e Piquete
    escalas = Escala.query.join(Bombeiro).filter(
        db.extract('month', Escala.data_inicio) == mes,
        db.extract('year', Escala.data_inicio) == ano,
        ~Escala.categoria.in_(['ECIN', 'ELAC', 'Piquete'])
    ).order_by(Escala.data_inicio.asc()).all()

    # Ordenação personalizada (sem data)
    categorias_ordem = ['Motorista', 'Socorrista', 'Centralista', 'EIP']
    prioridades = {
        'Motorista': ['Luís Matias', 'Jorge Pereira', 'José Soldado', 'José Seco', 'Pedro Fernandes',
                      'David Charrinho', 'Fábio Leirinha', 'Soeiro Mendes', 'Ana Marzia',
                      'Filipe Martins', 'Eric Nobre'],
        'Socorrista': ['José Rodrigues', 'Paulo Branquinho', 'Sabrina Fernandes'],
        'Centralista': ['Mariana Charrinho', 'Ruben Ramos', 'António Pequito'],
        'EIP': ['José Fernandes', 'João Mateus', 'Tiago Bizarro', 'João Carita', 'João Silva']
    }
    for cat in prioridades:
        prioridades[cat] = {nome: i for i, nome in enumerate(prioridades[cat])}

    def chave_ordenacao(esc):
        cat = esc.categoria if esc.categoria else 'Outros'
        ordem_cat = categorias_ordem.index(cat) if cat in categorias_ordem else len(categorias_ordem)
        nome = esc.bombeiro.nome.strip()
        if cat in prioridades:
            pos_nome = prioridades[cat].get(nome, len(prioridades[cat]))
            return (ordem_cat, pos_nome, nome)
        else:
            return (ordem_cat, esc.turno, nome)

    escalas = sorted(escalas, key=chave_ordenacao)

    # Construir estrutura ordenada
    from collections import OrderedDict
    estrutura = OrderedDict()
    for esc in escalas:
        bombeiro = esc.bombeiro
        cat = esc.categoria if esc.categoria else 'Outros'
        chave_bombeiro = (bombeiro.id, bombeiro.nome, bombeiro.mecanografico, cat)
        if chave_bombeiro not in estrutura:
            estrutura[chave_bombeiro] = {}
        try:
            num_turno = int(esc.turno.split('-')[0].strip())
        except (ValueError, IndexError):
            num_turno = 1
        dia = esc.data_inicio.day
        if dia not in estrutura[chave_bombeiro]:
            estrutura[chave_bombeiro][dia] = num_turno

    ultimo_dia = calendar.monthrange(ano, mes)[1]
    dias = list(range(1, ultimo_dia + 1))

    meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
    return render_template('imprimir_escala_mes.html',
                           estrutura=estrutura,
                           dias=dias,
                           mes=mes,
                           ano=ano,
                           meses=meses,
                           categorias_ordem=categorias_ordem,
                           feriados=feriados,
                           date=date)




@app.route('/escala/adicionar', methods=['POST'])
@login_required
def adicionar_escala():
    if current_user.tipo_user != 'Admin':
        flash('Apenas administração pode inserir escalas.', 'danger')
        return redirect(url_for('escala'))

    bombeiro_id = request.form['bombeiro_id']
    data_str = request.form['data_inicio']  # formato YYYY-MM-DD
    turno = request.form['turno']
    categoria = request.form.get('categoria', 'Bombeiro')
    funcao = request.form.get('funcao', '')

    # Converter a data para objeto date
    try:
        data = datetime.strptime(data_str, '%Y-%m-%d').date()
    except ValueError:
        flash('Data de início inválida.', 'danger')
        return redirect(url_for('escala'))

    # Determinar horas de início/fim com base no turno
    inicio_str, fim_str = turno_para_horas(turno)
    data_inicio = datetime.combine(data, datetime.strptime(inicio_str, '%H:%M').time())
    data_fim = datetime.combine(data, datetime.strptime(fim_str, '%H:%M').time())

    # Se a hora de fim for menor ou igual à de início, acrescenta 1 dia
    if data_fim <= data_inicio:
        data_fim += timedelta(days=1)

    nova = Escala(
        bombeiro_id=bombeiro_id,
        data_inicio=data_inicio,
        data_fim=data_fim,
        turno=turno,
        categoria=categoria,
        funcao=funcao
    )
    db.session.add(nova)
    db.session.commit()
    flash('Escala inserida.', 'success')
    return redirect(url_for('escala'))


def _parse_datetime(valor):
    """Tenta converter uma string de data/hora em datetime. Retorna None se falhar."""
    if not valor:
        return None
    formatos = [
        '%d/%m/%Y %H:%M',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%d/%m/%Y %H:%M:%S'
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(str(valor).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None

def _parse_data(valor):
    """Converte uma string de data (com ou sem hora) para um objeto date.
       Aceita vários formatos. Retorna None se falhar."""
    if not valor:
        return None
    s = str(valor).strip()
    # Se já veio como datetime (pode acontecer com openpyxl), extrai só a data
    if isinstance(valor, datetime):
        return valor.date()
    # Tenta formatos com hora (usa apenas a parte da data)
    formatos_com_hora = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%d/%m/%Y %H:%M:%S',
        '%d/%m/%Y %H:%M',
    ]
    for fmt in formatos_com_hora:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.date()
        except (ValueError, TypeError):
            continue
    # Tenta formatos apenas com data
    formatos_data = ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y']
    for fmt in formatos_data:
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            continue
    return None

@app.route('/escala/importar', methods=['POST'])
@login_required
def importar_escala():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito ao Comando/Admin.', 'danger')
        return redirect(url_for('escala'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('escala'))

    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '':
        flash('Ficheiro vazio.', 'warning')
        return redirect(url_for('escala'))

    if not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido. Use um ficheiro Excel (.xlsx).', 'danger')
        return redirect(url_for('escala'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler o ficheiro: {str(e)}', 'danger')
        return redirect(url_for('escala'))

    linhas_importadas = 0
    erros = []

    # Assumimos: Mecanográfico | Início | Fim | Turno | Categoria | Função (opcional)
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue

        try:
            mecanografico = str(row[0]).strip() if row[0] else None
            inicio_str = str(row[1]).strip() if row[1] else None
            fim_str = str(row[2]).strip() if row[2] else None
            turno = str(row[3]).strip() if len(row) > 3 and row[3] else None
            categoria = str(row[4]).strip() if len(row) > 4 and row[4] else 'Bombeiro'
            funcao = str(row[5]).strip() if len(row) > 5 and row[5] else None
        except IndexError:
            erros.append(f'Linha {row_num}: número insuficiente de colunas.')
            continue

        if not mecanografico or not inicio_str or not fim_str or not turno or not categoria:
            erros.append(f'Linha {row_num}: campos obrigatórios em falta (mecanográfico, início, fim, turno, categoria).')
            continue

        # Validar mecanográfico
        bombeiro = Bombeiro.query.filter_by(mecanografico=mecanografico).first()
        if not bombeiro:
            erros.append(f'Linha {row_num}: mecanográfico "{mecanografico}" não encontrado.')
            continue

        # Validar datas
        data_inicio = None
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
            try:
                data_inicio = datetime.strptime(inicio_str, fmt).date()
                break
            except ValueError:
                pass
        if not data_inicio:
            erros.append(f'Linha {row_num}: data de início inválida "{inicio_str}".')
            continue

        data_fim = None
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S'):
            try:
                data_fim = datetime.strptime(fim_str, fmt).date()
                break
            except ValueError:
                pass
        if not data_fim:
            erros.append(f'Linha {row_num}: data de fim inválida "{fim_str}".')
            continue

        # Validar turno
        turnos_validos = [
            '1 - 00h/08h', '2 - 08h/16h', '3 - 16h/24h', '4 - 11h/19h',
            '5 - 10h/18h', '6 - 07h/19h', '7 - 19h/07h', '8 - 08h/20h', '9 - 20h/08h'
        ]
        if turno not in turnos_validos:
            erros.append(f'Linha {row_num}: turno "{turno}" inválido.')
            continue

        # Validar categoria
        categorias_validas = ['Motorista', 'Socorrista', 'Centralista', 'EIP', 'ECIN', 'ELAC', 'Piquete', 'Bombeiro']
        if categoria not in categorias_validas:
            erros.append(f'Linha {row_num}: categoria "{categoria}" inválida.')
            continue

        # Criar a escala (sem tipo_bombeiro)
        nova = Escala(
            bombeiro_id=bombeiro.id,
            data_inicio=data_inicio,
            data_fim=data_fim,
            turno=turno,
            categoria=categoria,
            funcao=funcao if funcao else None
        )
        db.session.add(nova)
        linhas_importadas += 1

    db.session.commit()

    if erros:
        flash(f'{linhas_importadas} escala(s) importada(s) com sucesso. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} escala(s) importada(s) com sucesso!', 'success')

    return redirect(url_for('escala'))


@app.route('/escala/exportar')
@login_required
def exportar_escala():
    # Se quiser exportar apenas as escalas visíveis de acordo com os filtros atuais,
    # pode copiar a lógica da função 'escala' para construir a query.
    # Para simplificar, vamos exportar todas as escalas do mês atual, por exemplo.
    # Se preferir aplicar os mesmos filtros do utilizador, teria de passar os parâmetros.
    # Vou usar uma abordagem simples (exporta tudo ordenado).

    escalas = Escala.query.join(Bombeiro).order_by(Escala.data_inicio.asc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Escalas"

    cabecalhos = ['Mecanográfico', 'Nome', 'Início', 'Fim', 'Turno', 'Categoria', 'Função']
    ws.append(cabecalhos)
    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos) + 1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font

    for e in escalas:
        ws.append([
            e.bombeiro.mecanografico if e.bombeiro else '',
            e.bombeiro.nome if e.bombeiro else '',
            e.data_inicio.strftime('%d/%m/%Y %H:%M') if e.data_inicio else '',
            e.data_fim.strftime('%d/%m/%Y %H:%M') if e.data_fim else '',
            e.turno,
            e.categoria,
            e.funcao or ''
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='escalas.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ---------- Trocas de Serviço ----------
@app.route('/trocas', methods=['GET', 'POST'])
@login_required
def trocas():
    separador = request.args.get('tipo', 'assalariado')

    if request.method == 'POST':
        destino_id = request.form.get('destino_id', type=int)
        data_origem = datetime.strptime(request.form['data_origem'], '%Y-%m-%d').date()
        data_destino = datetime.strptime(request.form['data_destino'], '%Y-%m-%d').date()
        turno_origem = request.form.get('turno_origem', '')
        turno_destino = request.form.get('turno_destino', '')
        motivo = request.form.get('motivo', '')
        tipo_pedido = request.form.get('tipo_pedido', 'assalariado')   # campo oculto

        # Validação extra para ECINs
        if tipo_pedido == 'ecin':
            escalado = Escala.query.filter(
                Escala.bombeiro_id == current_user.id,
                func.date(Escala.data_inicio) <= data_origem,
                func.date(Escala.data_fim) >= data_origem,
                Escala.categoria.in_(['ECIN', 'ELAC'])
            ).first()
            if not escalado:
                flash('Não está escalado em ECIN/ELAC para esse dia.', 'danger')
                return redirect(url_for('trocas', tipo='ecin'))

        nova = TrocaServico(
            bombeiro_origem_id=current_user.id,
            bombeiro_destino_id=destino_id,
            data_origem=data_origem,
            data_destino=data_destino,
            turno_origem=turno_origem,
            turno_destino=turno_destino,
            motivo=motivo,
            estado='pendente_colega'
        )
        db.session.add(nova)
        db.session.commit()
        flash('Pedido de troca enviado.', 'success')
        return redirect(url_for('trocas', tipo=tipo_pedido))

    # --- GET ---
    query = TrocaServico.query

    if separador == 'assalariado':
        query = query.join(Bombeiro, TrocaServico.bombeiro_origem_id == Bombeiro.id)\
                     .filter(
                         Bombeiro.tipo_bombeiro == 'Profissional',
                         Bombeiro.posto.in_(['Motorista', 'Socorrista', 'Centralista'])
                     )
    else:  # ecin
        sub = db.session.query(TrocaServico.id)\
            .join(Escala, db.and_(
                Escala.bombeiro_id == TrocaServico.bombeiro_origem_id,
                func.date(Escala.data_inicio) <= TrocaServico.data_origem,
                func.date(Escala.data_fim) >= TrocaServico.data_origem,
                Escala.categoria.in_(['ECIN', 'ELAC'])
            )).subquery()
        query = query.filter(TrocaServico.id.in_(sub))

    if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Comando':
        pedidos = query.order_by(TrocaServico.data_pedido.desc()).all()
    else:
        pedidos = query.filter(
            (TrocaServico.bombeiro_origem_id == current_user.id) |
            (TrocaServico.bombeiro_destino_id == current_user.id)
        ).order_by(TrocaServico.data_pedido.desc()).all()

    bombeiros = Bombeiro.query.filter(Bombeiro.id != current_user.id, Bombeiro.ativo == True).all()
    return render_template('trocas.html', pedidos=pedidos, bombeiros=bombeiros, separador_atual=separador)

@app.route('/trocas/imprimir/<int:id>')
@login_required
def imprimir_troca(id):
    troca = TrocaServico.query.get_or_404(id)
    return render_template('imprimir_troca.html', troca=troca)


@app.route('/api/escala_usuario/<int:user_id>')
@login_required
def api_escala_usuario(user_id):
    ano = request.args.get('ano', type=int)
    mes = request.args.get('mes', type=int)
    if not ano or not mes:
        return {'erro': 'Parâmetros ano e mes obrigatórios'}, 400

    # Excluir escalas com observacao = 'dispensado' (e também 'troca de turno' se desejar que não apareçam,
    # mas vou manter 'troca de turno' visível, apenas ocultar 'dispensado')
    escalas = Escala.query.filter(
        Escala.bombeiro_id == user_id,
        db.extract('year', Escala.data_inicio) == ano,
        db.extract('month', Escala.data_inicio) == mes,
        (Escala.observacao != 'dispensado') | (Escala.observacao == None)
    ).order_by(Escala.data_inicio.asc()).all()

    result = {}
    for e in escalas:
        dia = e.data_inicio.day
        info = {
            'turno': e.turno,
            'categoria': e.categoria,
            'funcao': e.funcao or ''
        }
        if dia not in result:
            result[dia] = []
        result[dia].append(info)

    return result

@app.route('/escala/imprimir')
@login_required
def imprimir_escala():
    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int, default=datetime.now().year)
    categoria = request.args.get('categoria', '')
    mecanografico = request.args.get('mecanografico', '')

    query = Escala.query.join(Bombeiro)

    if mes:
        query = query.filter(db.extract('month', Escala.data_inicio) == mes)
        query = query.filter(db.extract('year', Escala.data_inicio) == ano)
    if categoria:
        query = query.filter(Escala.categoria == categoria)
    if mecanografico:
        query = query.filter(Bombeiro.mecanografico == mecanografico)

    if current_user.tipo_user != 'Admin':
        query = query.filter(Escala.bombeiro_id == current_user.id)

    escalas = query.order_by(Escala.data_inicio.asc()).all()

    # Determinar nome e mecanográfico
    nome_bombeiro = None
    mecanografico_bombeiro = None
    if mecanografico:
        bombeiro = Bombeiro.query.filter_by(mecanografico=mecanografico).first()
        if bombeiro:
            nome_bombeiro = bombeiro.nome
            mecanografico_bombeiro = bombeiro.mecanografico
    elif current_user.tipo_user != 'Admin':
        nome_bombeiro = current_user.nome
        mecanografico_bombeiro = current_user.mecanografico

    # Agrupar por categoria (ordem fixa)
    ordem_categorias = ['Motorista', 'Socorrista', 'Centralista', 'EIP', 'ECIN', 'ELAC', 'Piquete', 'Bombeiro']
    categorias_agrupadas = {}
    for cat in ordem_categorias:
        cats = [e for e in escalas if e.categoria == cat]
        if cats:
            categorias_agrupadas[cat] = cats

    return render_template('imprimir_escala.html',
                           categorias_agrupadas=categorias_agrupadas,
                           nome_bombeiro=nome_bombeiro,
                           mecanografico_bombeiro=mecanografico_bombeiro,
                           mes=mes,
                           ano=ano,
                           now=datetime.now())

# ---------- Aceitar/Recusar pelo colega (destino) ----------
@app.route('/trocas/aceitar/<int:id>')
@login_required
def aceitar_troca(id):
    troca = TrocaServico.query.get_or_404(id)
    # Só o destinatário pode aceitar
    if current_user.id != troca.bombeiro_destino_id:
        flash('Apenas o bombeiro recetor pode aceitar.', 'danger')
        return redirect(url_for('trocas'))
    if troca.estado != 'pendente_colega':
        flash('Este pedido já não está pendente.', 'warning')
        return redirect(url_for('trocas'))

    troca.estado = 'aceite_colega'
    db.session.commit()
    flash('Troca aceite. Aguarda aprovação do Comando.', 'success')
    return redirect(url_for('trocas'))


@app.route('/trocas/recusar/<int:id>')
@login_required
def recusar_troca(id):
    troca = TrocaServico.query.get_or_404(id)
    # Pode ser recusada pelo destinatário, Comando ou Admin
    permitido = (current_user.id == troca.bombeiro_destino_id) or \
                (current_user.tipo_user == 'Admin') or \
                (current_user.resp_departamento == 'Comando')
    if not permitido:
        flash('Sem permissão para recusar.', 'danger')
        return redirect(url_for('trocas'))
    if troca.estado in ['aprovada', 'recusada']:
        flash('Este pedido já foi finalizado.', 'warning')
        return redirect(url_for('trocas'))

    troca.estado = 'recusada'
    db.session.commit()
    flash('Troca recusada.', 'info')
    return redirect(url_for('trocas'))


@app.route('/trocas/aprovar/<int:id>')
@login_required
def aprovar_troca(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Apenas o Comando pode aprovar trocas.', 'danger')
        return redirect(url_for('trocas'))

    troca = TrocaServico.query.get_or_404(id)
    if troca.estado != 'aceite_colega':
        flash('A troca precisa de ser aceite pelo colega primeiro.', 'warning')
        return redirect(url_for('trocas'))

    escalas_origem = Escala.query.filter(
        Escala.bombeiro_id == troca.bombeiro_origem_id,
        Escala.data_inicio == troca.data_origem
    ).all()

    escalas_destino = Escala.query.filter(
        Escala.bombeiro_id == troca.bombeiro_destino_id,
        Escala.data_inicio == troca.data_destino
    ).all()

    for escala in escalas_origem:
        escala.bombeiro_id = troca.bombeiro_destino_id
        escala.observacao = 'troca de turno'
    for escala in escalas_destino:
        escala.bombeiro_id = troca.bombeiro_origem_id
        escala.observacao = 'troca de turno'

    troca.estado = 'aprovada'
    db.session.commit()

    total = len(escalas_origem) + len(escalas_destino)
    flash(f'Troca aprovada. {len(escalas_origem)} escala(s) na origem, {len(escalas_destino)} no destino atualizadas com observação.', 'success')
    return redirect(url_for('trocas'))





# ---------- Dispensas ----------
@app.route('/dispensas', methods=['GET', 'POST'])
@login_required
def dispensas():
    if request.method == 'POST':
        # Dados da dispensa
        data_inicio = datetime.strptime(request.form['data_inicio'], '%Y-%m-%d').date()
        data_fim = datetime.strptime(request.form['data_fim'], '%Y-%m-%d').date()
        motivo = request.form.get('motivo', '')
        # IDs dos créditos selecionados (vem como string separada por vírgulas)
        creditos_ids_str = request.form.get('creditos_selecionados', '')
        creditos_ids = [int(x.strip()) for x in creditos_ids_str.split(',') if x.strip().isdigit()] if creditos_ids_str else []

        nova = Dispensa(
            bombeiro_id=current_user.id,
            data_inicio=data_inicio,
            data_fim=data_fim,
            motivo=motivo
        )
        db.session.add(nova)
        db.session.commit()  # para obter o id

        # Se foram selecionados créditos, associa-os provisoriamente (só serão efetivados na aprovação)
        # Podemos guardar os ids numa coluna de texto ou criar tabela associativa, mas vou optar por já associar
        # e na aprovação apenas mudamos o estado. Mas o pedido é "quando aprovada pelo comandante, os dias selecionados na listbox mudam para Gozado".
        # Portanto, vamos já associar desde já, mas mantendo o estado 'Não Gozado' até aprovação.
        for cid in creditos_ids:
            credito = CreditoDispensa.query.get(cid)
            if credito and credito.bombeiro_id == current_user.id and credito.observacao == 'Não Gozado':
                credito.dispensa_id = nova.id
                # Ainda não mudamos a observacao; só na aprovação.
        db.session.commit()

        flash('Pedido de dispensa enviado.', 'success')
        return redirect(url_for('dispensas'))

    # GET – listagem de dispensas (já existente, com lógica admin/user)
    if current_user.tipo_user == 'Admin':
        dispensas_lista = Dispensa.query.order_by(Dispensa.id.desc()).all()
    else:
        dispensas_lista = Dispensa.query.filter_by(bombeiro_id=current_user.id).order_by(Dispensa.id.desc()).all()
    return render_template('dispensas.html', dispensas=dispensas_lista)

@app.route('/api/creditos_nao_gozados/<int:user_id>')
@login_required
def api_creditos_nao_gozados(user_id):
    creditos = CreditoDispensa.query.filter_by(
        bombeiro_id=user_id,
        observacao='Não Gozado'
    ).order_by(CreditoDispensa.data).all()
    resultado = [{
        'id': c.id,
        'data': c.data.strftime('%d/%m/%Y'),
        'descricao': c.descricao or '',
        'horas': c.horas
    } for c in creditos]
    return resultado


@app.route('/dispensas/aprovar/<int:id>')
@login_required
def aprovar_dispensa(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Apenas o Comando pode aprovar dispensas.', 'danger')
        return redirect(url_for('dispensas'))

    dispensa = Dispensa.query.get_or_404(id)
    if dispensa.aprovada:
        flash('Dispensa já aprovada.', 'info')
        return redirect(url_for('dispensas'))

    # Marcar dispensa como aprovada
    dispensa.aprovada = True

    # Atualizar créditos associados (Não Gozados) para Gozado
    creditos = CreditoDispensa.query.filter_by(dispensa_id=dispensa.id, observacao='Não Gozado').all()
    for cred in creditos:
        cred.observacao = 'Gozado'

    # Localizar escalas do bombeiro nas datas da dispensa e atualizar observação
    from datetime import timedelta
    dia_atual = dispensa.data_inicio
    while dia_atual <= dispensa.data_fim:
        escalas = Escala.query.filter(
            Escala.bombeiro_id == dispensa.bombeiro_id,
            Escala.data_inicio == dia_atual
        ).all()
        for e in escalas:
            e.observacao = 'dispensado'
        dia_atual += timedelta(days=1)

    db.session.commit()
    flash('Dispensa aprovada. Créditos atualizados e escala(s) marcada(s) como dispensado.', 'success')
    return redirect(url_for('dispensas'))

# ---------- Gestão de Créditos de Dispensa ----------
@app.route('/creditos')
@login_required
def listar_creditos():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito ao Comando/Admin.', 'danger')
        return redirect(url_for('dashboard'))

    bombeiro_id = request.args.get('bombeiro_id', type=int)
    query = CreditoDispensa.query
    if bombeiro_id:
        query = query.filter_by(bombeiro_id=bombeiro_id)

    # Ordenar por nome do bombeiro (fazendo join com Bombeiro)
    query = query.join(Bombeiro).order_by(Bombeiro.nome.asc())
    creditos = query.all()

    # Lista de bombeiros para o dropdown de filtro
    bombeiros = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()
    return render_template('creditos.html', creditos=creditos, bombeiros=bombeiros, bombeiro_id=bombeiro_id)



@app.route('/creditos/adicionar', methods=['POST'])
@login_required
def adicionar_credito():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_creditos'))

    bombeiro_id = request.form['bombeiro_id']
    data_str = request.form['data']
    descricao = request.form.get('descricao', '')
    horas = request.form.get('horas', 8, type=int)   # default 8

    data = datetime.strptime(data_str, '%Y-%m-%d').date()

    novo = CreditoDispensa(
        bombeiro_id=bombeiro_id,
        data=data,
        descricao=descricao,
        horas=horas,
        observacao='Em Análise'  # ← mude aqui (antes estava 'Não Gozado')
    )
    db.session.add(novo)
    db.session.commit()
    flash('Crédito adicionado.', 'success')
    return redirect(url_for('listar_creditos'))


@app.route('/creditos/apagar/<int:id>')
@login_required
def apagar_credito(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_creditos'))

    credito = CreditoDispensa.query.get_or_404(id)
    db.session.delete(credito)
    db.session.commit()
    flash('Crédito removido.', 'info')
    return redirect(url_for('listar_creditos'))


@app.route('/creditos/exportar')
@login_required
def exportar_creditos():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_creditos'))

    creditos = CreditoDispensa.query.order_by(CreditoDispensa.data.desc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Créditos Dispensa"

    # Cabeçalhos
    cabecalhos = ['ID', 'Mecanográfico', 'Nome', 'Data', 'Descrição', 'Horas', 'Estado']
    ws.append(cabecalhos)

    # Estilo do cabeçalho
    from openpyxl.styles import Font, PatternFill
    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    # Dados
    for c in creditos:
        ws.append([
            c.id,
            c.bombeiro.mecanografico,
            c.bombeiro.nome,
            c.data.strftime('%d/%m/%Y') if c.data else '',
            c.descricao or '',
            c.horas,
            c.observacao
        ])

    # Ajustar largura das colunas
    col_widths = [5, 15, 30, 12, 30, 8, 12]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    from flask import send_file
    return send_file(
        output,
        as_attachment=True,
        download_name='creditos_dispensa.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/creditos/importar', methods=['POST'])
@login_required
def importar_creditos():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_creditos'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('listar_creditos'))

    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '':
        flash('Ficheiro vazio.', 'warning')
        return redirect(url_for('listar_creditos'))

    if not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido. Use .xlsx.', 'danger')
        return redirect(url_for('listar_creditos'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {e}', 'danger')
        return redirect(url_for('listar_creditos'))

    linhas_importadas = 0
    erros = []

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue

        try:
            mecanografico_val = row[0]
            data_val = row[1]
            descricao_val = row[2] if len(row) > 2 else None
            horas_val = row[3] if len(row) > 3 else 8
        except IndexError:
            erros.append(f'Linha {row_num}: colunas insuficientes.')
            continue

        mecanografico = str(mecanografico_val).strip() if mecanografico_val is not None else ''
        if not mecanografico:
            erros.append(f'Linha {row_num}: mecanográfico em falta.')
            continue

        bombeiro = Bombeiro.query.filter_by(mecanografico=mecanografico).first()
        if not bombeiro:
            erros.append(f'Linha {row_num}: mecanográfico "{mecanografico}" não encontrado.')
            continue

        # Parse da data usando a função _parse_data (já existe no app.py)
        data_credito = _parse_data(data_val)
        if not data_credito:
            erros.append(f'Linha {row_num}: data inválida "{data_val}".')
            continue

        descricao = str(descricao_val).strip() if descricao_val is not None else ''
        try:
            horas = int(horas_val) if horas_val is not None else 8
        except (ValueError, TypeError):
            horas = 8

        novo = CreditoDispensa(
            bombeiro_id=bombeiro.id,
            data=data_credito,
            descricao=descricao,
            horas=horas,
            observacao='Não Gozado'
        )
        db.session.add(novo)
        linhas_importadas += 1

    db.session.commit()

    if erros:
        flash(f'{linhas_importadas} crédito(s) importado(s). {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} crédito(s) importado(s) com sucesso!', 'success')

    return redirect(url_for('listar_creditos'))

# ---------- Créditos users ----------

@app.route('/meus-creditos', methods=['GET', 'POST'])
@login_required
def meus_creditos():
    # Filtro por estado (Gozado / Não Gozado)
    estado = request.args.get('estado', '')  # 'Gozado', 'Não Gozado' ou vazio (todos)

    if request.method == 'POST':
        data_str = request.form['data']
        descricao = request.form.get('descricao', '')
        horas = request.form.get('horas', 8, type=int)

        try:
            data = datetime.strptime(data_str, '%Y-%m-%d').date()
        except ValueError:
            flash('Data inválida.', 'danger')
            return redirect(url_for('meus_creditos'))

        novo = CreditoDispensa(
            bombeiro_id=current_user.id,
            data=data,
            descricao=descricao,
            horas=horas,
            observacao='Em Análise'  # ← agora fica pendente
        )
        db.session.add(novo)
        db.session.commit()
        flash('Crédito registado e em análise.', 'success')
        return redirect(url_for('meus_creditos'))

    # Construir query para os créditos do utilizador atual
    query = CreditoDispensa.query.filter_by(bombeiro_id=current_user.id)
    if estado:
        query = query.filter_by(observacao=estado)

    creditos = query.order_by(CreditoDispensa.data.desc()).all()

    # Estados disponíveis para o filtro
    estados = ['Não Gozado', 'Gozado']

    return render_template('meus_creditos.html',
                           creditos=creditos,
                           estados=estados,
                           estado_atual=estado)


# ----------  Aprovação Admin Créditos  ----------
@app.route('/creditos/aprovar/<int:id>')
@login_required
def aprovar_credito(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_creditos'))

    credito = CreditoDispensa.query.get_or_404(id)
    if credito.observacao == 'Em Análise':
        credito.observacao = 'Não Gozado'
        db.session.commit()
        flash('Crédito aprovado (estado: Não Gozado).', 'success')
    else:
        flash('Este crédito não está em análise.', 'warning')

    return redirect(url_for('listar_creditos'))


# ---------- Disponibilidades ----------
@app.route('/disponibilidades', methods=['GET', 'POST'])
@login_required
def disponibilidades():
    if request.method == 'POST':
        # Recolher as datas enviadas (separadas por vírgula)
        datas_str = request.form.get('datas', '')
        categoria = request.form.get('categoria', '')
        turno_extra = request.form.get('turno_extra', '')

        if not datas_str:
            flash('Selecione pelo menos uma data.', 'warning')
            return redirect(url_for('disponibilidades'))

        datas = [d.strip() for d in datas_str.split(',') if d.strip()]
        for data_str in datas:
            try:
                data = datetime.strptime(data_str, '%Y-%m-%d').date()
            except ValueError:
                continue
            # Verificar se já existe uma disponibilidade idêntica
            existente = Disponibilidade.query.filter_by(
                bombeiro_id=current_user.id,
                data=data,
                turno_extra=turno_extra,
                categoria=categoria
            ).first()
            if existente:
                continue
            nova = Disponibilidade(
                bombeiro_id=current_user.id,
                data=data,
                turno_extra=turno_extra,
                categoria=categoria,
                confirmada=False
            )
            db.session.add(nova)
        db.session.commit()
        flash('Disponibilidade(s) registada(s).', 'success')
        return redirect(url_for('disponibilidades'))

    # GET – listagem com filtros
    bombeiro_id = request.args.get('bombeiro_id', type=int)
    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)

    query = Disponibilidade.query

    # Apenas Admin/Comando podem ver todos; os restantes ficam limitados ao seu ID
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        query = query.filter_by(bombeiro_id=current_user.id)
    elif bombeiro_id:
        query = query.filter_by(bombeiro_id=bombeiro_id)

    if mes and ano:
        query = query.filter(
            db.extract('month', Disponibilidade.data) == mes,
            db.extract('year', Disponibilidade.data) == ano
        )
    elif mes:
        query = query.filter(db.extract('month', Disponibilidade.data) == mes)
    elif ano:
        query = query.filter(db.extract('year', Disponibilidade.data) == ano)

    lista = query.order_by(Disponibilidade.data.desc()).all()

    # Lista de bombeiros ativos (apenas para Admin/Comando)
    bombeiros_ativos = []
    if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Comando':
        bombeiros_ativos = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()

    now = date.today()
    return render_template('disponibilidades.html',
                           disponibilidades=lista,
                           bombeiros=bombeiros_ativos,
                           now=now,
                           bombeiro_id=bombeiro_id,
                           mes=mes,
                           ano=ano)

@app.route('/disponibilidades/apagar', methods=['POST'])
@login_required
def apagar_disponibilidades_mes():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('disponibilidades'))

    mes = request.form.get('mes', type=int)
    ano = request.form.get('ano', type=int)
    bombeiro_id = request.form.get('bombeiro_id', type=int)  # opcional

    if not mes or not ano:
        flash('Mês e ano são obrigatórios para apagar.', 'warning')
        return redirect(url_for('disponibilidades'))

    query = Disponibilidade.query.filter(
        db.extract('month', Disponibilidade.data) == mes,
        db.extract('year', Disponibilidade.data) == ano
    )
    if bombeiro_id:
        query = query.filter_by(bombeiro_id=bombeiro_id)

    count = query.count()
    query.delete()
    db.session.commit()
    flash(f'{count} disponibilidade(s) apagada(s) (mês {mes}/{ano}).', 'success')
    return redirect(url_for('disponibilidades', mes=mes, ano=ano, bombeiro_id=bombeiro_id))


@app.route('/disponibilidades/aprovar', methods=['POST'])
@login_required
def aprovar_disponibilidades():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('disponibilidades'))

    ids_str = request.form.get('ids', '')
    if not ids_str:
        flash('Nenhuma disponibilidade selecionada.', 'warning')
        return redirect(url_for('disponibilidades'))

    ids = [int(i) for i in ids_str.split(',') if i.strip().isdigit()]
    for id in ids:
        disp = Disponibilidade.query.get(id)
        if disp and not disp.confirmada:
            disp.confirmada = True

            # Criar ECIN automaticamente
            existente = Ecin.query.filter_by(
                bombeiro_id=disp.bombeiro_id,
                data=disp.data,
                turno=disp.turno_extra
            ).first()
            if not existente:
                novo_ecin = Ecin(
                    bombeiro_id=disp.bombeiro_id,
                    data=disp.data,
                    turno=disp.turno_extra,
                    estado='Pendente'
                )
                db.session.add(novo_ecin)

    db.session.commit()
    flash(f'{len(ids)} disponibilidade(s) aprovada(s) e registada(s) em ECINS.', 'success')
    return redirect(url_for('disponibilidades'))


@app.route('/disponibilidades/imprimir')
@login_required
def imprimir_disponibilidades():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('disponibilidades'))

    bombeiro_id = request.args.get('bombeiro_id', type=int)
    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)

    query = Disponibilidade.query
    if bombeiro_id:
        query = query.filter_by(bombeiro_id=bombeiro_id)
    if mes and ano:
        query = query.filter(db.extract('month', Disponibilidade.data) == mes,
                             db.extract('year', Disponibilidade.data) == ano)
    elif ano:
        query = query.filter(db.extract('year', Disponibilidade.data) == ano)

    lista = query.order_by(Disponibilidade.data.asc()).all()
    return render_template('imprimir_disponibilidades.html', disponibilidades=lista)



# ---------- Confirmar Disponibiidade ----------
@app.route('/disponibilidades/confirmar', methods=['POST'])
@login_required
def confirmar_disponibilidade():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('disponibilidades'))

    id = request.form.get('id', type=int)
    if not id:
        flash('ID em falta.', 'warning')
        return redirect(url_for('disponibilidades'))

    disp = Disponibilidade.query.get_or_404(id)
    disp.confirmada = True

    # Criar automaticamente um registo ECIN, se ainda não existir
    existente = Ecin.query.filter_by(
        bombeiro_id=disp.bombeiro_id,
        data=disp.data,
        turno=disp.turno_extra
    ).first()
    if not existente:
        novo_ecin = Ecin(
            bombeiro_id=disp.bombeiro_id,
            data=disp.data,
            turno=disp.turno_extra,
            estado='Pendente'
        )
        db.session.add(novo_ecin)

    db.session.commit()
    flash('Disponibilidade confirmada e registada em ECINS.', 'success')
    return redirect(url_for('disponibilidades'))



#_________Exportar disponibilidades----
@app.route('/disponibilidades/exportar')
@login_required
def exportar_disponibilidades():
    if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Comando':
        lista = Disponibilidade.query.order_by(Disponibilidade.data.desc()).all()
    else:
        lista = Disponibilidade.query.filter_by(bombeiro_id=current_user.id)\
                                     .order_by(Disponibilidade.data.desc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Disponibilidades"

    cabecalhos = ['Bombeiro', 'Data', 'Turno', 'Confirmada']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for d in lista:
        ws.append([
            d.bombeiro.nome,
            d.data.strftime('%d/%m/%Y') if d.data else '',
            d.turno_extra,
            'Sim' if d.confirmada else 'Não'
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='disponibilidades.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------- Ecins ----------
@app.route('/ecins')
@login_required
def listar_ecins():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    ordem = request.args.get('ordem', 'data')
    mec = request.args.get('mec', '').strip()

    query = Ecin.query
    if mec:
        query = query.join(Bombeiro).filter(Bombeiro.mecanografico.ilike(f'%{mec}%'))

    if ordem == 'nome':
        query = query.join(Bombeiro).order_by(Bombeiro.nome.asc(), Ecin.data.asc())
    else:
        query = query.order_by(Ecin.data.asc())

    registos = query.all()

    # Para o modal "Novo Registo" vamos usar todos os bombeiros ativos
    bombeiros_ativos = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()

    # Turnos específicos para ECIN/ELAC
    turnos_ecinelac = ['07h/19h', '19h/07h']

    return render_template('ecins.html',
                           registos=registos,
                           bombeiros_ativos=bombeiros_ativos,
                           turnos=turnos_ecinelac,
                           agora=date.today(),
                           ordem_atual=ordem,
                           mec_pesquisa=mec)

# ---------- Adicionar Ecins ----------
@app.route('/ecins/adicionar', methods=['POST'])
@login_required
def adicionar_ecin():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_ecins'))

    bombeiro_id = request.form['bombeiro_id']
    data_str = request.form['data']
    turno = request.form['turno']
    categoria = request.form['categoria']   # 'ECIN' ou 'ELAC'

    try:
        data = datetime.strptime(data_str, '%Y-%m-%d').date()
    except ValueError:
        flash('Data inválida.', 'danger')
        return redirect(url_for('listar_ecins'))

    novo = Ecin(
        bombeiro_id=bombeiro_id,
        data=data,
        turno=turno,
        categoria=categoria,
        estado='Pendente'
    )
    db.session.add(novo)
    db.session.commit()
    flash('Registo ECIN criado.', 'success')
    return redirect(url_for('listar_ecins'))

# ----------Escalar Ecins ----------

@app.route('/ecins/escalar/<int:id>')
@login_required
def escalar_ecin(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_ecins'))

    ecin = Ecin.query.get_or_404(id)
    # Permitir alteração mesmo que já não esteja 'Pendente'
    funcao_cod = request.args.get('funcao', 'X')

    mapeamento = {
        'M':  {'funcao': 'Motorista', 'categoria': 'ECIN', 'estado': 'Motorista ECIN'},
        'C':  {'funcao': 'Chefe',     'categoria': 'ECIN', 'estado': 'Chefe ECIN'},
        'G':  {'funcao': 'Guarnição', 'categoria': 'ECIN', 'estado': 'Guarnição ECIN'},
        'Me': {'funcao': 'Motorista', 'categoria': 'ELAC', 'estado': 'Motorista ELAC'},
        'Ce': {'funcao': 'Chefe',     'categoria': 'ELAC', 'estado': 'Chefe ELAC'},
        'X':  {'estado': 'Não Escalado'}
    }

    if funcao_cod not in mapeamento:
        flash('Opção inválida.', 'danger')
        return redirect(url_for('listar_ecins'))

    dados = mapeamento[funcao_cod]

    # Se já tinha sido escalado anteriormente, remover a escala antiga
    if ecin.estado not in ['Pendente', 'Não Escalado']:
        escalas_antigas = Escala.query.filter_by(
            bombeiro_id=ecin.bombeiro_id,
            data_inicio=datetime.combine(ecin.data, datetime.strptime('07:00', '%H:%M').time())  # pode ser genérico
        ).filter(
            Escala.turno == ecin.turno,
            Escala.categoria.in_(['ECIN', 'ELAC'])
        ).all()
        for esc in escalas_antigas:
            db.session.delete(esc)

    if funcao_cod == 'X':
        ecin.estado = dados['estado']
        ecin.funcao = None
        ecin.categoria = None
        db.session.commit()
        flash('Marcado como Não Escalado.', 'info')
        return redirect(url_for('listar_ecins'))

    # Criar nova escala
    try:
        partes = ecin.turno.split('-')[1].strip().split('/')
        inicio_str = partes[0].replace('h', ':00')
        fim_str = partes[1].replace('h', ':00')
    except Exception:
        inicio_str, fim_str = '08:00', '20:00'

    data = ecin.data
    inicio_time = datetime.strptime(inicio_str, '%H:%M').time()
    fim_time = datetime.strptime(fim_str, '%H:%M').time()
    data_inicio = datetime.combine(data, inicio_time)
    data_fim = datetime.combine(data, fim_time)
    if data_fim <= data_inicio:
        data_fim += timedelta(days=1)

    nova_escala = Escala(
        bombeiro_id=ecin.bombeiro_id,
        data_inicio=data_inicio,
        data_fim=data_fim,
        turno=ecin.turno,
        categoria=dados['categoria'],
        funcao=dados['funcao']
    )
    db.session.add(nova_escala)

    ecin.funcao = dados['funcao']
    ecin.categoria = dados['categoria']
    ecin.estado = dados['estado']
    db.session.commit()
    flash(f'{ecin.bombeiro.nome} escalado como {dados["estado"]}.', 'success')
    return redirect(url_for('listar_ecins'))

#----------Imprimir Disponibilidade ECINS----------------
@app.route('/ecins/imprimir')
@login_required
def imprimir_ecins():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)
    if not mes or not ano:
        hoje = date.today()
        mes = hoje.month
        ano = hoje.year

    # Filtrar ECINS do mês/ano, ordenados por data
    ecins = Ecin.query.filter(
        db.extract('month', Ecin.data) == mes,
        db.extract('year', Ecin.data) == ano
    ).order_by(Ecin.data).all()

    # Agrupar por dia e turno
    from collections import defaultdict
    escala = defaultdict(lambda: defaultdict(list))
    for ec in ecins:
        dia_str = ec.data.strftime('%d/%m/%Y')
        turno_str = ec.turno
        # Mapear turnos conhecidos: "07h/19h" e "19h/07h"
        if turno_str in ['07h/19h', '19h/07h']:
            escala[dia_str][turno_str].append(ec.bombeiro.nome)
        else:
            # Se não for um dos turnos padrão, colocar em "Outro"
            escala[dia_str]['Outro'].append(ec.bombeiro.nome)

    # Ordenar dias cronologicamente
    dias_ordenados = sorted(escala.keys(), key=lambda d: datetime.strptime(d, '%d/%m/%Y'))

    meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
                   'Novembro', 'Dezembro']

    return render_template('imprimir_ecins.html',
                           escala=escala,
                           dias=dias_ordenados,
                           mes=mes,
                           ano=ano,
                           meses=meses_nomes)

@app.route('/ecins/modificar/<int:id>')
@login_required
def modificar_ecin(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('listar_ecins'))

    ecin = Ecin.query.get_or_404(id)
    if ecin.estado == 'Pendente':
        flash('Registo já está pendente.', 'info')
        return redirect(url_for('listar_ecins'))

    # Remover a escala correspondente (se existir)
    # Pode procurar na tabela Escala pelo bombeiro, data, turno, categoria, função...
    escala = Escala.query.filter_by(
        bombeiro_id=ecin.bombeiro_id,
        data_inicio=datetime.combine(ecin.data, datetime.strptime('00:00', '%H:%M').time()),  # pode ser aproximado
        # Melhor: usar os campos que tem no ecin para identificar a escala.
    ).first()
    # Melhor: guardar no ecin o id da escala criada (opcional). Vamos fazer de forma mais simples:
    # Apagar qualquer escala do mesmo bombeiro, data, turno, categoria = ecin.categoria e funcao = ecin.funcao
    if ecin.categoria and ecin.funcao:
        escalas = Escala.query.filter_by(
            bombeiro_id=ecin.bombeiro_id,
            turno=ecin.turno,
            categoria=ecin.categoria,
            funcao=ecin.funcao
        ).filter(
            func.date(Escala.data_inicio) == ecin.data
        ).all()
        for escala in escalas:
            db.session.delete(escala)
    elif ecin.categoria:
        # se não tem função, pode ter ido para Não Escalado, sem escala
        pass

    ecin.estado = 'Pendente'
    ecin.funcao = None
    ecin.categoria = None
    db.session.commit()
    flash('Registo modificado. Escolha novamente.', 'info')
    return redirect(url_for('listar_ecins'))

# ---------- Imprimir Escala ECin----------

@app.route('/ecins/imprimir-escala-ecin')
@login_required
def imprimir_escala_ecin():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)
    if not mes or not ano:
        hoje = date.today()
        mes = hoje.month
        ano = hoje.year

    # Filtrar apenas ECINs com categoria 'ECIN' e estados Motorista/Chefe/Guarnição
    ecins = Ecin.query.filter(
        Ecin.categoria == 'ECIN',
        Ecin.estado.in_(['Motorista ECIN', 'Chefe ECIN', 'Guarnição ECIN']),
        db.extract('month', Ecin.data) == mes,
        db.extract('year', Ecin.data) == ano
    ).order_by(Ecin.data).all()

    from collections import defaultdict
    escala = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))  # dia -> turno -> funcao -> nomes

    for ec in ecins:
        dia_str = ec.data.strftime('%d/%m/%Y')
        turno_str = ec.turno
        if turno_str in ['07h/19h', '19h/07h']:
            funcao = ec.funcao if ec.funcao else 'Outro'
            escala[dia_str][turno_str][funcao].append({
                'nome': ec.bombeiro.nome,
                'mecanografico': ec.bombeiro.mecanografico,
                'posto': ec.bombeiro.posto
            })

    dias_ordenados = sorted(escala.keys(), key=lambda d: datetime.strptime(d, '%d/%m/%Y'))
    meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
                   'Novembro', 'Dezembro']

    return render_template('imprimir_ecin_escala.html',
                           escala=escala,
                           dias=dias_ordenados,
                           mes=mes,
                           ano=ano,
                           meses=meses_nomes)


@app.route('/ecins/imprimir-escala-elac')
@login_required
def imprimir_escala_elac():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Secretaria']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    mes = request.args.get('mes', type=int)
    ano = request.args.get('ano', type=int)
    if not mes or not ano:
        hoje = date.today()
        mes = hoje.month
        ano = hoje.year

    # Filtrar apenas ELACs com categoria 'ELAC' e estados Motorista/Chefe
    elacs = Ecin.query.filter(
        Ecin.categoria == 'ELAC',
        Ecin.estado.in_(['Motorista ELAC', 'Chefe ELAC']),
        db.extract('month', Ecin.data) == mes,
        db.extract('year', Ecin.data) == ano
    ).order_by(Ecin.data).all()

    from collections import defaultdict
    escala = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))  # dia -> turno -> funcao -> lista de bombeiros

    for ec in elacs:
        dia_str = ec.data.strftime('%d/%m/%Y')
        turno_str = ec.turno
        if turno_str in ['07h/19h', '19h/07h']:
            funcao = ec.funcao if ec.funcao else 'Outro'
            escala[dia_str][turno_str][funcao].append({
                'nome': ec.bombeiro.nome,
                'mecanografico': ec.bombeiro.mecanografico,
                'posto': ec.bombeiro.posto
            })

    dias_ordenados = sorted(escala.keys(), key=lambda d: datetime.strptime(d, '%d/%m/%Y'))
    meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
                   'Novembro', 'Dezembro']

    return render_template('imprimir_elac_escala.html',
                           escala=escala,
                           dias=dias_ordenados,
                           mes=mes,
                           ano=ano,
                           meses=meses_nomes)

# ---------- Bombeiro Perfil ----------

@app.route('/perfil', methods=['GET', 'POST'])
@login_required
def perfil():
    bombeiro = current_user  # já é o objeto do utilizador logado

    if request.method == 'POST':
        novo_email = request.form['email'].strip().lower()
        novo_telemovel = request.form.get('telemovel', '').strip()
        nova_password = request.form.get('password', '')

        # Verificar duplicação de email (exceto o próprio)
        conflito_email = Bombeiro.query.filter(
            Bombeiro.id != bombeiro.id,
            Bombeiro.email == novo_email
        ).first()
        if conflito_email:
            flash('Este email já está a ser usado por outro bombeiro.', 'warning')
            return redirect(url_for('perfil'))

        # Verificar duplicação de telemóvel (se preenchido)
        if novo_telemovel:
            conflito_tele = Bombeiro.query.filter(
                Bombeiro.id != bombeiro.id,
                Bombeiro.telemovel == novo_telemovel
            ).first()
            if conflito_tele:
                flash('Este número de telemóvel já está associado a outro bombeiro.', 'warning')
                return redirect(url_for('perfil'))

        # Atualizar campos
        bombeiro.email = novo_email
        bombeiro.telemovel = novo_telemovel if novo_telemovel else None

        # Só altera a password se foi preenchida
        if nova_password:
            bombeiro.password_hash = generate_password_hash(nova_password)

        db.session.commit()
        flash('Perfil atualizado com sucesso.', 'success')
        return redirect(url_for('perfil'))

    # GET → mostra o formulário pré‑preenchido
    return render_template('perfil.html')



# ---------- Checklist ----------
@app.route('/checklist', methods=['GET', 'POST'])
@login_required
def checklist():
    if request.method == 'POST':
        viatura_id = request.form['viatura_id']
        itens = request.form['itens']
        observacoes = request.form.get('observacoes', '')
        nova = Checklist(
            viatura_id=viatura_id,
            bombeiro_id=current_user.id,
            itens_verificados=itens,
            observacoes=observacoes
        )
        db.session.add(nova)
        db.session.commit()
        flash('Checklist registado.', 'success')
        return redirect(url_for('checklist'))

    checklists_lista = Checklist.query.order_by(Checklist.data.desc()).all()
    viaturas = Viatura.query.all()
    return render_template('checklist.html', checklists=checklists_lista, viaturas=viaturas)


# ---------- Fardamento ----------
@app.route('/fardamento', methods=['GET', 'POST'])
@login_required
def fardamento():
    if request.method == 'POST':
        tipo = request.form['tipo']
        nome = request.form['nome']
        tamanho = request.form['tamanho']
        motivo = request.form['motivo']
        descricao_motivo = request.form.get('descricao_motivo', '')
        stock_id = request.form.get('stock_id', type=int)

        # Obter descrição automaticamente do stock
        descricao = ''
        if stock_id:
            item_stock = StockFardamento.query.get(stock_id)
            if item_stock:
                descricao = item_stock.descricao or ''

        novo = Fardamento(
            bombeiro_id=current_user.id,
            tipo=tipo,
            nome=nome,
            descricao=descricao,
            tamanho=tamanho,
            motivo=motivo,
            descricao_motivo=descricao_motivo,
            stock_id=stock_id,
            estado='Pedido'
        )
        db.session.add(novo)
        db.session.commit()
        flash('Pedido de fardamento registado.', 'success')
        return redirect(url_for('fardamento'))

    # GET
    if current_user.tipo_user == 'Admin' or current_user.resp_departamento in ['Comando', 'Fardamento']:
        pedidos = Fardamento.query.order_by(Fardamento.data_registo.desc()).all()
    else:
        pedidos = Fardamento.query.filter_by(bombeiro_id=current_user.id)\
                                  .order_by(Fardamento.data_registo.desc()).all()

    # Listas para dropdowns dinâmicos
    tipos_stock = db.session.query(StockFardamento.tipo).distinct().all()
    tipos_stock = [t[0] for t in tipos_stock if t[0]]
    todos_itens_stock = StockFardamento.query.order_by(StockFardamento.nome).all()

    return render_template('fardamento.html',
                           pedidos=pedidos,
                           tipos_stock=tipos_stock,
                           todos_itens_stock=todos_itens_stock)


# ---------- Fardamento p/tipo ----------
@app.route('/fardamento/nomes-por-tipo/<tipo>')
@login_required
def nomes_por_tipo(tipo):
    itens = StockFardamento.query.filter_by(tipo=tipo).order_by(StockFardamento.nome).all()
    result = [{
        'id': i.id,
        'nome': i.nome,
        'descricao': i.descricao or '',
        'tamanho': i.tamanho or '',
        'stock': i.stock                  # ← adicionado
    } for i in itens]
    from flask import jsonify
    return jsonify(result)

# ---------- Editar Fardamento ----------

@app.route('/fardamento/editar/<int:id>', methods=['POST'])
@login_required
def editar_fardamento(id):
    pedido = Fardamento.query.get_or_404(id)

    # Permissão: Admin/Comando/Fardamento ou o próprio bombeiro dono do pedido
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento'] and current_user.id != pedido.bombeiro_id:
        flash('Sem permissão.', 'danger')
        return redirect(url_for('fardamento'))

    # Atualizar campos básicos (só o dono ou admin pode mudar o pedido em estado 'Pedido')
    if pedido.estado == 'Pedido' or current_user.tipo_user == 'Admin' or current_user.resp_departamento in ['Comando', 'Fardamento']:
        pedido.tipo = request.form.get('tipo', pedido.tipo)
        pedido.nome = request.form.get('nome', pedido.nome)
        pedido.tamanho = request.form.get('tamanho', pedido.tamanho)
        pedido.motivo = request.form.get('motivo', pedido.motivo)
        pedido.descricao_motivo = request.form.get('descricao_motivo', '') if pedido.motivo in ['Troca', 'Abatido'] else ''
        stock_id = request.form.get('stock_id', type=int)
        if stock_id:
            pedido.stock_id = stock_id
            item = StockFardamento.query.get(stock_id)
            if item:
                pedido.descricao = item.descricao or ''

    # Campos de aprovação (só Admin/Comando/Fardamento)
    if current_user.tipo_user == 'Admin' or current_user.resp_departamento in ['Comando', 'Fardamento']:
        # Responsável (apenas Admin ou Fardamento)
        if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Fardamento':
            pedido.responsavel = request.form.get('responsavel') == 'on'
        # Comando
        if current_user.tipo_user == 'Admin' or current_user.resp_departamento == 'Comando':
            pedido.comando = request.form.get('comando') == 'on'

        # Entregue (só se responsável e comando estiverem marcados)
        if pedido.responsavel and pedido.comando:
            pedido.entregue = request.form.get('entregue') == 'on'
            if pedido.entregue:
                pedido.data_entrega = datetime.utcnow()
                pedido.estado = 'Concluido'
                # Atualizar stock
                if pedido.stock_id:
                    item_stock = StockFardamento.query.get(pedido.stock_id)
                    if item_stock and item_stock.stock > 0:
                        item_stock.stock -= 1
        else:
            pedido.entregue = False

        # Estado (pode ser alterado manualmente)
        novo_estado = request.form.get('estado')
        if novo_estado in ['Pedido', 'Análise', 'Concluido']:
            pedido.estado = novo_estado

    db.session.commit()
    flash('Pedido atualizado.', 'success')
    return redirect(url_for('fardamento'))

# ---------- Apagar Fardamento ----------

@app.route('/fardamento/apagar/<int:id>')
@login_required
def apagar_fardamento(id):
    pedido = Fardamento.query.get_or_404(id)
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento'] and current_user.id != pedido.bombeiro_id:
        flash('Sem permissão.', 'danger')
        return redirect(url_for('fardamento'))
    db.session.delete(pedido)
    db.session.commit()
    flash('Pedido removido.', 'info')
    return redirect(url_for('fardamento'))


# ---------- Exportar Fardamento ----------
@app.route('/fardamento/exportar')
@login_required
def exportar_fardamento():
    if current_user.tipo_user == 'Admin' or current_user.resp_departamento in ['Comando', 'Fardamento']:
        pedidos = Fardamento.query.order_by(Fardamento.data_registo.desc()).all()
    else:
        pedidos = Fardamento.query.filter_by(bombeiro_id=current_user.id)\
                                  .order_by(Fardamento.data_registo.desc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fardamento"

    cabecalhos = ['ID', 'Data', 'Mecanográfico', 'Bombeiro', 'Tipo', 'Nome', 'Tamanho',
                  'Motivo', 'Descrição Motivo', 'Estado', 'Resp.', 'Comando', 'Entregue', 'Data Entrega']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for p in pedidos:
        ws.append([
            p.id,
            p.data_registo.strftime('%d/%m/%Y %H:%M') if p.data_registo else '',
            p.bombeiro.mecanografico,
            p.bombeiro.nome,
            p.tipo,
            p.nome,
            p.tamanho,
            p.motivo,
            p.descricao_motivo or '',
            p.estado,
            'Sim' if p.responsavel else 'Não',
            'Sim' if p.comando else 'Não',
            'Sim' if p.entregue else 'Não',
            p.data_entrega.strftime('%d/%m/%Y %H:%M') if p.data_entrega else ''
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='fardamento.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ---------- Importar Fardamento ----------
@app.route('/fardamento/importar', methods=['POST'])
@login_required
def importar_fardamento():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('fardamento'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('fardamento'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('fardamento'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('fardamento'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            mecanografico = str(row[2]).strip() if len(row) > 2 and row[2] else None
            tipo = str(row[4]).strip() if len(row) > 4 and row[4] else ''
            nome = str(row[5]).strip() if len(row) > 5 and row[5] else ''
            tamanho = str(row[6]).strip() if len(row) > 6 and row[6] else ''
            motivo = str(row[7]).strip() if len(row) > 7 and row[7] else 'Novo'
            estado = str(row[9]).strip() if len(row) > 9 and row[9] else 'Pedido'
        except Exception:
            erros.append(f'Linha {row_num}: dados inválidos.')
            continue

        bombeiro = Bombeiro.query.filter_by(mecanografico=mecanografico).first() if mecanografico else None
        if not bombeiro:
            erros.append(f'Linha {row_num}: mecanográfico não encontrado.')
            continue

        # Encontrar stock_id pelo nome e tipo (simplificação)
        stock_item = StockFardamento.query.filter_by(nome=nome, tipo=tipo).first()
        stock_id = stock_item.id if stock_item else None

        novo = Fardamento(
            bombeiro_id=bombeiro.id,
            tipo=tipo,
            nome=nome,
            tamanho=tamanho,
            motivo=motivo,
            estado=estado,
            stock_id=stock_id
        )
        db.session.add(novo)
        linhas_importadas += 1

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importados. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} pedidos importados com sucesso.', 'success')
    return redirect(url_for('fardamento'))


#-----------Stock Fardamento--------------
@app.route('/stock-fardamento', methods=['GET', 'POST'])
@login_required
def stock_fardamento():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        tipo = request.form.get('tipo', 'Outro')
        nome = request.form['nome']
        descricao = request.form.get('descricao', '')
        tamanho = request.form.get('tamanho', '')
        stock = request.form.get('stock', 0, type=int)
        novo = StockFardamento(
            tipo=tipo,
            nome=nome,
            descricao=descricao,
            tamanho=tamanho,
            stock=stock
        )
        db.session.add(novo)
        db.session.commit()
        flash('Item de fardamento adicionado ao stock.', 'success')
        return redirect(url_for('stock_fardamento'))

    itens = StockFardamento.query.order_by(StockFardamento.nome).all()
    return render_template('stock_fardamento.html', itens=itens)

#-----------Editar Stock Fardamento--------------

@app.route('/stock-fardamento/editar/<int:id>', methods=['POST'])
@login_required
def editar_stock_fardamento(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_fardamento'))

    item = StockFardamento.query.get_or_404(id)
    item.tipo = request.form.get('tipo', 'Outro')
    item.nome = request.form['nome']
    item.descricao = request.form.get('descricao', '')
    item.tamanho = request.form.get('tamanho', '')
    item.stock = request.form.get('stock', 0, type=int)
    db.session.commit()
    flash('Item atualizado.', 'success')
    return redirect(url_for('stock_fardamento'))

#-----------Editar Stock Fardamento--------------
@app.route('/stock-fardamento/apagar/<int:id>')
@login_required
def apagar_stock_fardamento(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_fardamento'))

    item = StockFardamento.query.get_or_404(id)
    db.session.delete(item)
    db.session.commit()
    flash('Item removido.', 'info')
    return redirect(url_for('stock_fardamento'))

#-----------exportar Stock Fardamento--------------
@app.route('/stock-fardamento/exportar')
@login_required
def exportar_stock_fardamento():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_fardamento'))

    itens = StockFardamento.query.order_by(StockFardamento.nome).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock Fardamento"

    cabecalhos = ['Tipo','Nome', 'Descrição', 'Tamanho', 'Stock']
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for i in itens:
        ws.append([i.tipo, i.nome, i.descricao or '', i.tamanho or '', i.stock])

    col_widths = [30, 40, 12, 22, 10]
    for j, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='stock_fardamento.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


#-----------Importar Stock Fardamento--------------
@app.route('/stock-fardamento/importar', methods=['POST'])
@login_required
def importar_stock_fardamento():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_fardamento'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('stock_fardamento'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('stock_fardamento'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('stock_fardamento'))

    linhas_importadas = 0
    erros = []

    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue

        try:
            tipo = str(row[0]).strip() if len(row) > 0 and row[0] else 'Outro'
            nome = str(row[1]).strip() if row[1] else ''
            descricao = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            tamanho = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            stock_str = str(row[4]).strip() if len(row) > 4 and row[4] else '0'
            stock = int(stock_str)
        except (ValueError, TypeError) as e:
            erros.append(f'Linha {row_num}: valor inválido - {str(e)}')
            continue
        except Exception as e:
            erros.append(f'Linha {row_num}: erro inesperado - {str(e)}')
            continue

        if not nome:
            erros.append(f'Linha {row_num}: nome obrigatório.')
            continue

        novo = StockFardamento(
            tipo=tipo,
            nome=nome,
            descricao=descricao,
            tamanho=tamanho,
            stock=stock
        )
        db.session.add(novo)
        linhas_importadas += 1

    db.session.commit()

    if erros:
        flash(f'{linhas_importadas} importados. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} itens importados com sucesso!', 'success')
    return redirect(url_for('stock_fardamento'))

#-----------Imprimir Stock Fardamento--------------
@app.route('/stock-fardamento/imprimir')
@login_required
def imprimir_stock_fardamento():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_fardamento'))

    itens = StockFardamento.query.order_by(StockFardamento.nome).all()
    return render_template('imprimir_stock_fardamento.html', itens=itens)

#------------Farmaneto Atribuido-------------
@app.route('/fardamento-atribuido', methods=['GET', 'POST'])
@login_required
def fardamento_atribuido():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        bombeiro_id = request.form['bombeiro_id']
        tipo = request.form['tipo']
        nome = request.form['nome']
        tamanho = request.form['tamanho']
        data_entrega = datetime.strptime(request.form['data_entrega'], '%Y-%m-%d').date()

        novo = FardamentoAtribuido(
            bombeiro_id=bombeiro_id,
            tipo=tipo,
            nome=nome,
            tamanho=tamanho,
            data_entrega=data_entrega,
            estado='Entregue'
        )
        db.session.add(novo)
        db.session.commit()
        flash('Fardamento atribuído.', 'success')
        return redirect(url_for('fardamento_atribuido'))

    bombeiro_id = request.args.get('bombeiro_id', type=int)
    query = FardamentoAtribuido.query
    if bombeiro_id:
        query = query.filter_by(bombeiro_id=bombeiro_id)
    atribuicoes = query.order_by(FardamentoAtribuido.data_entrega.desc()).all()
    bombeiros = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()

    return render_template('fardamento_atribuido.html',
                           atribuicoes=atribuicoes,
                           bombeiros=bombeiros,
                           bombeiro_id=bombeiro_id)


#------------Devoluçã Farmaneto-------------

@app.route('/fardamento-atribuido/devolver/<int:id>', methods=['POST'])
@login_required
def devolver_fardamento_atribuido(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Fardamento']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))
    item = FardamentoAtribuido.query.get_or_404(id)
    item.data_devolucao = date.today()
    item.estado = 'Devolvido'
    db.session.commit()
    flash('Devolução registada.', 'success')
    return redirect(url_for('fardamento_atribuido'))


#----------- Stock Farmácia--------------

@app.route('/stock-farmacia', methods=['GET', 'POST'])
@login_required
def stock_farmacia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        categoria = request.form['categoria']
        nome = request.form['nome']
        tamanho = request.form.get('tamanho', '')
        stock = request.form.get('stock', 0, type=int)

        novo = StockFarmacia(
            categoria=categoria,
            nome=nome,
            tamanho=tamanho,
            stock=stock,
            data_atualizacao=datetime.utcnow()
        )
        db.session.add(novo)
        db.session.commit()
        flash('Produto de farmácia adicionado.', 'success')
        return redirect(url_for('stock_farmacia'))

    itens = StockFarmacia.query.order_by(StockFarmacia.categoria, StockFarmacia.nome).all()
    categorias = CategoriaFarmacia.query.order_by(CategoriaFarmacia.nome).all()
    return render_template('stock_farmacia.html', itens=itens, categorias=categorias)


@app.route('/stock-farmacia/editar/<int:id>', methods=['POST'])
@login_required
def editar_stock_farmacia(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_farmacia'))

    item = StockFarmacia.query.get_or_404(id)
    item.categoria = request.form['categoria']
    item.nome = request.form['nome']
    item.tamanho = request.form.get('tamanho', '')
    novo_stock = request.form.get('stock', 0, type=int)
    if novo_stock != item.stock:
        item.data_atualizacao = datetime.utcnow()
    item.stock = novo_stock
    db.session.commit()
    flash('Produto atualizado.', 'success')
    return redirect(url_for('stock_farmacia'))


@app.route('/stock-farmacia/apagar/<int:id>')
@login_required
def apagar_stock_farmacia(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_farmacia'))
    item = StockFarmacia.query.get_or_404(id)
    db.session.delete(item)
    db.session.commit()
    flash('Produto removido.', 'info')
    return redirect(url_for('stock_farmacia'))



@app.route('/categorias-farmacia', methods=['GET', 'POST'])
@login_required
def categorias_farmacia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        nome = request.form['nome'].strip()
        checklist = request.form.get('checklist') == 'on'   # ← captura a checkbox

        if not nome:
            flash('Nome da categoria obrigatório.', 'warning')
        elif CategoriaFarmacia.query.filter_by(nome=nome).first():
            flash('Categoria já existe.', 'warning')
        else:
            nova = CategoriaFarmacia(nome=nome, checklist=checklist)
            db.session.add(nova)
            db.session.commit()
            flash('Categoria criada.', 'success')
        return redirect(url_for('categorias_farmacia'))

    # GET – listar categorias
    categorias = CategoriaFarmacia.query.order_by(CategoriaFarmacia.nome).all()
    return render_template('categorias_farmacia.html', categorias=categorias)


#----------- Exportar Stock Farmácia--------------
@app.route('/stock-farmacia/exportar')
@login_required
def exportar_stock_farmacia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_farmacia'))

    itens = StockFarmacia.query.order_by(StockFarmacia.categoria, StockFarmacia.nome).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock Farmacia"

    cabecalhos = ['ID', 'Categoria', 'Nome', 'Tamanho', 'Stock', 'Atualização']
    ws.append(cabecalhos)
    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)
    for col in range(1, len(cabecalhos)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font

    for i in itens:
        ws.append([i.id, i.categoria, i.nome, i.tamanho or '', i.stock,
                   i.data_atualizacao.strftime('%d/%m/%Y %H:%M') if i.data_atualizacao else ''])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='stock_farmacia.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

#----------- Importar Stock Farmácia--------------
@app.route('/stock-farmacia/importar', methods=['POST'])
@login_required
def importar_stock_farmacia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_farmacia'))

    if 'ficheiro' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('stock_farmacia'))
    ficheiro = request.files['ficheiro']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido.', 'danger')
        return redirect(url_for('stock_farmacia'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
        ws = wb.active
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('stock_farmacia'))

    linhas_importadas = 0
    erros = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(cell is None for cell in row):
            continue
        try:
            categoria = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            nome = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            tamanho = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            stock = int(row[4]) if len(row) > 4 and row[4] else 0
        except Exception:
            erros.append(f'Linha {row_num}: dados inválidos.')
            continue
        if not nome:
            erros.append(f'Linha {row_num}: nome obrigatório.')
            continue
        novo = StockFarmacia(categoria=categoria, nome=nome, tamanho=tamanho, stock=stock,
                             data_atualizacao=datetime.utcnow())
        db.session.add(novo)
        linhas_importadas += 1

    db.session.commit()
    if erros:
        flash(f'{linhas_importadas} importados. {len(erros)} erro(s): ' + '; '.join(erros), 'warning')
    else:
        flash(f'{linhas_importadas} produtos importados com sucesso.', 'success')
    return redirect(url_for('stock_farmacia'))

#------------Imprimir Stock Farmacia----------
@app.route('/stock-farmacia/imprimir')
@login_required
def imprimir_stock_farmacia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_farmacia'))

    itens = StockFarmacia.query.order_by(StockFarmacia.categoria, StockFarmacia.nome).all()
    return render_template('imprimir_stock_farmacia.html', itens=itens)



#----------- Categorias Farmácia--------------
@app.route('/categorias-farmacia/apagar/<int:id>')
@login_required
def apagar_categoria_farmacia(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('categorias_farmacia'))
    cat = CategoriaFarmacia.query.get_or_404(id)
    db.session.delete(cat)
    db.session.commit()
    flash('Categoria removida.', 'info')
    return redirect(url_for('categorias_farmacia'))


#----------- CheckList Ambulancia--------------

@app.route('/checklist-ambulancia', methods=['GET', 'POST'])
@login_required
def checklist_ambulancia():
    if request.method == 'POST':
        viatura_id = request.form['viatura_id']
        # Criar a checklist com o utilizador autenticado
        nova = ChecklistAmbulancia(viatura_id=viatura_id, bombeiro_id=current_user.id)
        db.session.add(nova)
        db.session.flush()

        # Produtos das categorias com checklist ativo
        categorias_check = CategoriaFarmacia.query.filter_by(checklist=True).all()
        ids_categorias = [c.id for c in categorias_check]
        produtos_disponiveis = StockFarmacia.query.filter(
            db.func.lower(StockFarmacia.categoria).in_([c.nome.lower() for c in categorias_check])
        ).all() if categorias_check else []

        ids_selecionados = []
        for produto in produtos_disponiveis:
            if request.form.get(f'prod_{produto.id}') == 'on':
                ids_selecionados.append(produto.id)

        for pid in ids_selecionados:
            item = ChecklistAmbulanciaItem(checklist_id=nova.id, produto_id=pid, quantidade=0)
            db.session.add(item)

        if not ids_selecionados:
            db.session.rollback()
            flash('Selecione pelo menos um produto.', 'warning')
            return redirect(url_for('checklist_ambulancia'))

        db.session.commit()
        return redirect(url_for('preencher_quantidades', checklist_id=nova.id))

    # GET – listagem com filtro opcional por viatura
    viatura_id = request.args.get('viatura_id', type=int)
    query = ChecklistAmbulancia.query
    if viatura_id:
        query = query.filter_by(viatura_id=viatura_id)
    checklists = query.order_by(ChecklistAmbulancia.data_hora.desc()).all()

    viaturas = Viatura.query.filter(Viatura.tipo.in_(['ABSC', 'ABTD', 'ABTM'])).order_by(Viatura.matricula).all()

    # Dados para o modal de criação
    categorias = CategoriaFarmacia.query.filter_by(checklist=True).order_by(CategoriaFarmacia.nome).all()
    produtos_por_categoria = []
    for cat in categorias:
        produtos = StockFarmacia.query.filter(
            db.func.lower(StockFarmacia.categoria) == cat.nome.lower()
        ).order_by(StockFarmacia.nome).all()
        if produtos:
            produtos_por_categoria.append((cat, produtos))

    return render_template('checklist_ambulancia.html',
                           checklists=checklists,
                           viaturas=viaturas,
                           viatura_selecionada=viatura_id,
                           produtos_por_categoria=produtos_por_categoria)



@app.route('/checklist-ambulancia/quantidades/<int:checklist_id>', methods=['GET', 'POST'])
@login_required
def preencher_quantidades(checklist_id):
    checklist = ChecklistAmbulancia.query.get_or_404(checklist_id)
    if checklist.finalizado:
        flash('Checklist já finalizada.', 'warning')
        return redirect(url_for('checklist_ambulancia'))

    if request.method == 'POST':
        # Atualizar quantidades nos itens da checklist e criar StockAmbulancia (pendente)
        for item in checklist.itens:
            qtd_str = request.form.get(f'qtd_{item.id}', '0')
            try:
                qtd = int(qtd_str)
            except ValueError:
                qtd = 0
            item.quantidade = qtd

            if qtd > 0:
                # Criar pedido no Stock Ambulância (pendente, sem abater)
                reposicao = StockAmbulancia(
                    ambulancia_id=checklist.viatura_id,
                    produto_id=item.produto_id,
                    quantidade=qtd,
                    solicitante_id=current_user.id,
                    responsavel_id=None,
                    checklist_id=checklist.id,
                    confirmado=False
                )
                db.session.add(reposicao)

        checklist.finalizado = True
        db.session.commit()
        flash('Quantidades registadas e pedidos enviados para Stock Ambulância.', 'success')
        return redirect(url_for('checklist_ambulancia'))


    # GET – mostrar formulário com os itens selecionados
    # Responsáveis: bombeiros cujo departamento é 'Farmácia'
    responsaveis = Bombeiro.query.filter_by(resp_departamento='Farmácia', ativo=True).all()
    return render_template('preencher_quantidades.html', checklist=checklist, responsaveis=responsaveis)

@app.route('/checklist-ambulancia/imprimir-lista')
@login_required
def imprimir_lista_checklists():
    viatura_id = request.args.get('viatura_id', type=int)
    query = ChecklistAmbulancia.query
    if viatura_id:
        query = query.filter_by(viatura_id=viatura_id)
    checklists = query.order_by(ChecklistAmbulancia.data_hora.desc()).all()
    return render_template('imprimir_lista_checklists.html', checklists=checklists)


@app.route('/checklist-ambulancia/imprimir/<int:checklist_id>')
@login_required
def imprimir_checklist(checklist_id):
    checklist = ChecklistAmbulancia.query.get_or_404(checklist_id)
    return render_template('imprimir_checklist_ambulancia.html', checklist=checklist)


#----------- Stock Ambulancia--------------

@app.route('/stock-ambulancia')
@login_required
def stock_ambulancia():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia', 'Socorrista']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    reposicoes = StockAmbulancia.query.order_by(StockAmbulancia.data.desc()).all()
    return render_template('stock_ambulancia.html', reposicoes=reposicoes)

@app.route('/stock-ambulancia/confirmar/<int:id>')
@login_required
def confirmar_reposicao(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento not in ['Comando', 'Farmacia']:
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('stock_ambulancia'))

    reposicao = StockAmbulancia.query.get_or_404(id)
    if reposicao.confirmado:
        flash('Reposição já confirmada.', 'warning')
        return redirect(url_for('stock_ambulancia'))

    # Verificar stock e abater
    produto = reposicao.produto_stock
    if produto.stock < reposicao.quantidade:
        flash(f'Stock insuficiente de "{produto.nome}" (disponível: {produto.stock}).', 'danger')
        return redirect(url_for('stock_ambulancia'))

    produto.stock -= reposicao.quantidade
    produto.data_atualizacao = datetime.utcnow()

    reposicao.confirmado = True
    reposicao.responsavel_id = current_user.id
    db.session.commit()
    flash('Reposição confirmada e stock atualizado.', 'success')
    return redirect(url_for('stock_ambulancia'))


#_____________________Central_____________

@app.route('/central')
@login_required
def central():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        flash('Acesso restrito ao Departamento Central.', 'danger')
        return redirect(url_for('dashboard'))

    hoje = date.today()
    daqui_7_dias = hoje + timedelta(days=7)

    # Totais de escalados hoje por categoria
    categorias = ['Motorista', 'Socorrista', 'Centralista', 'EIP', 'ECIN', 'ELAC', 'Piquete', 'Bombeiro']
    totais = {}
    for cat in categorias:
        count = Escala.query.filter(
            func.date(Escala.data_inicio) <= hoje,
            func.date(Escala.data_fim) >= hoje,
            Escala.categoria == cat
        ).distinct(Escala.bombeiro_id).count()
        totais[cat] = count

    viaturas_inop = Viatura.query.filter(
        func.lower(Viatura.estado) == 'inoperacional'
    ).order_by(Viatura.matricula).all()
    notas_7dias = Nota.query.filter(
        Nota.data_evento >= hoje,
        Nota.data_evento <= daqui_7_dias
    ).order_by(Nota.data_evento.asc()).all()

    bombeiros = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()
    departamentos = list(set(b.resp_departamento for b in bombeiros if b.resp_departamento))

    return render_template('central.html',
                           hoje=hoje,
                           totais=totais,
                           viaturas_inop=viaturas_inop,
                           notas_7dias=notas_7dias,
                           bombeiros=bombeiros,
                           departamentos=departamentos)



@app.route('/central/dia/<string:data_str>')
@login_required
def central_dia(data_str):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        return jsonify({'erro': 'Acesso restrito'}), 403

    try:
        data = datetime.strptime(data_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'erro': 'Formato de data inválido'}), 400

    categoria = request.args.get('categoria', '')  # novo parâmetro

    html = gerar_html_dia(data, categoria)
    return jsonify({'html': html})

@app.route('/central/totais/<string:data_str>')
@login_required
def central_totais_dia(data_str):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        return jsonify({'erro': 'Acesso restrito'}), 403
    try:
        data = datetime.strptime(data_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'erro': 'Data inválida'}), 400

    categorias = ['Motorista', 'Socorrista', 'Centralista', 'EIP', 'ECIN', 'ELAC', 'Piquete', 'Bombeiro']
    totais = {}
    for cat in categorias:
        count = Escala.query.filter(
            func.date(Escala.data_inicio) <= data,
            func.date(Escala.data_fim) >= data,
            Escala.categoria == cat
        ).distinct(Escala.bombeiro_id).count()
        totais[cat] = count
    return jsonify(totais)


@app.route('/central/notas', methods=['POST'])
@login_required
def central_notas():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    descricao = request.form['descricao']
    data_evento_str = request.form.get('data_evento')
    data_evento = datetime.strptime(data_evento_str, '%Y-%m-%d').date() if data_evento_str else None

    nota = Nota(
        criador_id=current_user.id,
        descricao=descricao,
        data_evento=data_evento
    )
    db.session.add(nota)
    db.session.commit()
    flash('Nota registada com sucesso.', 'success')
    return redirect(url_for('central', tab='notas'))

@app.route('/central/notas/editar/<int:id>', methods=['POST'])
@login_required
def central_editar_nota(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('central', tab='notas'))
    nota = Nota.query.get_or_404(id)
    nota.descricao = request.form['descricao']
    data_evento_str = request.form.get('data_evento')
    nota.data_evento = datetime.strptime(data_evento_str, '%Y-%m-%d').date() if data_evento_str else None
    db.session.commit()
    flash('Nota atualizada.', 'success')
    return redirect(url_for('central', tab='notas'))

@app.route('/central/notas/apagar/<int:id>')
@login_required
def central_apagar_nota(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('central', tab='notas'))
    nota = Nota.query.get_or_404(id)
    db.session.delete(nota)
    db.session.commit()
    flash('Nota removida.', 'info')
    return redirect(url_for('central', tab='notas'))



@app.route('/central/atividade_mes/<int:ano>/<int:mes>')
@login_required
def central_atividade_mes(ano, mes):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Central':
        return jsonify({'erro': 'Acesso restrito'}), 403

    hoje = date.today()
    primeiro_dia = date(ano, mes, 1)
    ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])

    atividade = {}
    for dia in range(1, ultimo_dia.day + 1):
        d = date(ano, mes, dia)
        tem_escalas = Escala.query.filter(
            func.date(Escala.data_inicio) <= d,
            func.date(Escala.data_fim) >= d
        ).first() is not None
        tem_trocas = TrocaServico.query.filter(
            (TrocaServico.data_origem == d) | (TrocaServico.data_destino == d),
            TrocaServico.estado == 'aprovada'
        ).first() is not None
        tem_dispensas = Dispensa.query.filter(
            Dispensa.data_inicio <= d,
            Dispensa.data_fim >= d,
            Dispensa.aprovada == True
        ).first() is not None

        atividade[str(d)] = {
            'escalas': tem_escalas,
            'trocas': tem_trocas,
            'dispensas': tem_dispensas
        }

    return jsonify(atividade)




#----------------Secção Correio---------------
@app.route('/correio')
@login_required
def correio():
    # Caixa de entrada: mensagens onde o current_user é destinatário OU pertence ao departamento destino
    caixa_entrada = MensagemCorreio.query.filter(
        (
                (MensagemCorreio.destinatario_id == current_user.id) |
                (MensagemCorreio.departamento == current_user.resp_departamento)
        ),
        MensagemCorreio.remetente_id != current_user.id,  # ← exclui as enviadas pelo próprio
        MensagemCorreio.apagada_destinatario == False
    ).order_by(MensagemCorreio.data_envio.desc()).all()

    # Mensagens enviadas pelo current_user (histórico)
    enviadas = MensagemCorreio.query.filter(
        MensagemCorreio.remetente_id == current_user.id,
        MensagemCorreio.apagada_remetente == False
    ).order_by(MensagemCorreio.data_envio.desc()).all()

    # Para o formulário de envio
    bombeiros = Bombeiro.query.filter_by(ativo=True).order_by(Bombeiro.nome).all()
    departamentos = list(set(b.resp_departamento for b in bombeiros if b.resp_departamento))

    return render_template('correio.html',
                           caixa_entrada=caixa_entrada,
                           enviadas=enviadas,
                           bombeiros=bombeiros,
                           departamentos=departamentos)


@app.route('/correio/enviar', methods=['POST'])
@login_required
def correio_enviar():
    destinatario_tipo = request.form.get('destinatario_tipo')  # 'bombeiro' ou 'departamento'
    destinatario_id = request.form.get('destinatario_id', type=int)
    departamento = request.form.get('departamento') if destinatario_tipo == 'departamento' else None
    assunto = request.form.get('assunto', 'Sem assunto')
    corpo = request.form['corpo']

    nova = MensagemCorreio(
        remetente_id=current_user.id,
        destinatario_id=destinatario_id if destinatario_tipo == 'bombeiro' else None,
        departamento=departamento,
        assunto=assunto,
        corpo=corpo
    )
    db.session.add(nova)
    db.session.commit()
    flash('Mensagem enviada com sucesso.', 'success')
    return redirect(url_for('correio'))

@app.route('/correio/lida/<int:id>')
@login_required
def correio_marcar_lida(id):
    msg = MensagemCorreio.query.get_or_404(id)
    if msg.destinatario_id == current_user.id or msg.departamento == current_user.resp_departamento:
        msg.lida = True
        db.session.commit()
    return redirect(url_for('correio'))

@app.route('/correio/apagar/<int:id>')
@login_required
def correio_apagar(id):
    msg = MensagemCorreio.query.get_or_404(id)
    if msg.remetente_id == current_user.id:
        msg.apagada_remetente = True
    elif msg.destinatario_id == current_user.id or msg.departamento == current_user.resp_departamento:
        msg.apagada_destinatario = True
    else:
        flash('Acesso negado.', 'danger')
        return redirect(url_for('correio'))
    db.session.commit()
    flash('Mensagem removida.', 'info')
    return redirect(url_for('correio'))


@app.context_processor
def inject_novas_mensagens():
    if current_user.is_authenticated:
        naolidas = MensagemCorreio.query.filter(
            (
                (MensagemCorreio.destinatario_id == current_user.id) |
                (MensagemCorreio.departamento == current_user.resp_departamento)
            ),
            MensagemCorreio.remetente_id != current_user.id,
            MensagemCorreio.lida == False,
            MensagemCorreio.apagada_destinatario == False
        ).count()
    else:
        naolidas = 0
    return dict(msgs_naolidas=naolidas)

#------------------Importar 03-05-26----------------
# ========== FUNÇÕES AUXILIARES DE IMPORTAÇÃO ==========

def _importar_linha_bombeiros(row, row_num):
    numero = str(row[0]).strip() if row[0] else ''
    mecanografico = str(row[1]).strip() if len(row) > 1 and row[1] else ''
    nome = str(row[2]).strip() if len(row) > 2 and row[2] else ''
    nomecompleto = str(row[3]).strip() if len(row) > 3 and row[3] else ''
    email = str(row[4]).strip().lower() if len(row) > 4 and row[4] else ''
    telemovel = str(row[5]).strip() if len(row) > 5 and row[5] else None
    posto = str(row[6]).strip() if len(row) > 6 and row[6] else ''
    tipo_bombeiro = str(row[7]).strip() if len(row) > 7 and row[7] else 'Voluntário'
    departamento = str(row[8]).strip() if len(row) > 8 and row[8] else None
    tipo_user = str(row[9]).strip() if len(row) > 9 and row[9] else 'User'
    ativo = str(row[10]).strip().lower() == 'sim' if len(row) > 10 and row[10] else True
    password_hash = str(row[11]).strip() if len(row) > 11 and row[11] else generate_password_hash('123456')

    if not numero or not mecanografico or not nome or not email:
        return "Campos obrigatórios em falta (Nº Interno, Mecanográfico, Nome, Email)"
    if Bombeiro.query.filter(
        (Bombeiro.numero_interno == numero) |
        (Bombeiro.mecanografico == mecanografico) |
        (Bombeiro.email == email)
    ).first():
        return "Duplicado"
    b = Bombeiro(
        numero_interno=numero,
        mecanografico=mecanografico,
        nome=nome,
        nomecompleto=nomecompleto if nomecompleto else None,
        email=email,
        telemovel=telemovel if telemovel and telemovel != '' else None,
        posto=posto,
        tipo_bombeiro=tipo_bombeiro,
        resp_departamento=departamento if departamento and departamento != '' else None,
        tipo_user=tipo_user,
        ativo=ativo,
        password_hash=password_hash
    )
    db.session.add(b)
    db.session.flush()
    return None


def _importar_linha_viaturas(row, row_num):
    matricula = str(row[0]).strip() if row[0] else ''
    if not matricula:
        return "Matrícula obrigatória"
    if Viatura.query.filter_by(matricula=matricula).first():
        return "Matrícula já existe"
    tipo = str(row[1]).strip() if len(row) > 1 else ''
    nomenclatura = str(row[2]).strip() if len(row) > 2 else ''
    marca = str(row[3]).strip() if len(row) > 3 else ''
    modelo = str(row[4]).strip() if len(row) > 4 else ''
    ano = int(row[5]) if len(row) > 5 and row[5] else 0
    estado = str(row[6]).strip().lower() if len(row) > 6 else 'operacional'
    v = Viatura(matricula=matricula, tipo=tipo, nomenclatura=nomenclatura,
                marca=marca, modelo=modelo, ano=ano, estado=estado)
    db.session.add(v)
    db.session.flush()
    return None


def _importar_linha_categorias_farmacia(row, row_num):
    nome = str(row[0]).strip() if row[0] else ''
    if not nome:
        return "Nome da categoria obrigatório"
    if CategoriaFarmacia.query.filter_by(nome=nome).first():
        return "Categoria já existe"
    checklist = str(row[1]).strip().lower() == 'sim' if len(row) > 1 else False
    c = CategoriaFarmacia(nome=nome, checklist=checklist)
    db.session.add(c)
    db.session.flush()
    return None


def _importar_linha_stock_farmacia(row, row_num):
    categoria = str(row[0]).strip() if row[0] else ''
    nome = str(row[1]).strip() if len(row) > 1 else ''
    if not nome:
        return "Nome do produto obrigatório"
    tamanho = str(row[2]).strip() if len(row) > 2 else ''
    stock = int(row[3]) if len(row) > 3 and row[3] else 0
    s = StockFarmacia(categoria=categoria, nome=nome, tamanho=tamanho, stock=stock,
                      data_atualizacao=datetime.utcnow())
    db.session.add(s)
    db.session.flush()
    return None


def _importar_linha_stock_fardamento(row, row_num):
    nome = str(row[0]).strip() if row[0] else ''
    descricao = str(row[1]).strip() if len(row) > 1 else ''
    tamanho = str(row[2]).strip() if len(row) > 2 else ''
    tipo = str(row[3]).strip() if len(row) > 3 else ''
    stock = int(row[4]) if len(row) > 4 and row[4] else 0
    s = StockFardamento(nome=nome, descricao=descricao, tamanho=tamanho, tipo=tipo, stock=stock)
    db.session.add(s)
    db.session.flush()
    return None


def _importar_linha_disponibilidades(row, row_num):
    mec = str(row[0]).strip() if row[0] else None
    data_str = str(row[1]).strip() if len(row) > 1 else None
    turno_extra = str(row[2]).strip() if len(row) > 2 else ''
    categoria = str(row[3]).strip() if len(row) > 3 else ''
    confirmada = str(row[4]).strip().lower() == 'sim' if len(row) > 4 else False
    data = _parse_data(data_str) if data_str else None
    if not mec or not data:
        return "Mecanográfico ou data obrigatórios"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    d = Disponibilidade(bombeiro_id=bombeiro.id, data=data, turno_extra=turno_extra,
                        categoria=categoria, confirmada=confirmada)
    db.session.add(d)
    db.session.flush()
    return None


def _importar_linha_escalas(row, row_num):
    mec = str(row[0]).strip() if row[0] else None
    inicio_str = str(row[1]).strip() if len(row) > 1 else None
    fim_str = str(row[2]).strip() if len(row) > 2 else None
    turno = str(row[3]).strip() if len(row) > 3 else ''
    categoria = str(row[4]).strip() if len(row) > 4 else 'Bombeiro'
    funcao = str(row[5]).strip() if len(row) > 5 else None
    inicio = _parse_datetime(inicio_str)
    fim = _parse_datetime(fim_str)
    if not mec or not inicio or not fim:
        return "Dados de escala incompletos"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    e = Escala(bombeiro_id=bombeiro.id, data_inicio=inicio, data_fim=fim,
               turno=turno, categoria=categoria,
               funcao=funcao if funcao and funcao != '' else None)
    db.session.add(e)
    db.session.flush()
    return None


def _importar_linha_avarias(row, row_num):
    codigo = str(row[0]).strip() if row[0] else None
    matricula = str(row[1]).strip() if len(row) > 1 else None
    descricao = str(row[2]).strip() if len(row) > 2 else ''
    reportador_mec = str(row[3]).strip() if len(row) > 3 else None
    kms = int(row[4]) if len(row) > 4 and row[4] else None
    resp_oficina = str(row[5]).strip().lower() == 'sim' if len(row) > 5 else False
    comando = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False
    estado = str(row[7]).strip() if len(row) > 7 else 'Pendente'
    data_str = str(row[8]).strip() if len(row) > 8 else None

    if not descricao or not matricula:
        return "Descrição e matrícula obrigatórias"
    viatura = Viatura.query.filter_by(matricula=matricula).first()
    if not viatura:
        return f"Viatura {matricula} não encontrada"
    reportador = Bombeiro.query.filter_by(mecanografico=reportador_mec).first() if reportador_mec else None
    data_reporte = _parse_datetime(data_str) or datetime.utcnow()

    a = Avaria(codigo=codigo, viatura_id=viatura.id,
               descricao=descricao, reportado_por=reportador.id if reportador else 1,
               kms=kms, responsavel_oficina=resp_oficina, comando_verificado=comando,
               estado=estado, data_reporte=data_reporte)
    db.session.add(a)
    db.session.flush()
    return None


def _importar_linha_trocas(row, row_num):
    mec_origem = str(row[0]).strip() if row[0] else None
    mec_destino = str(row[1]).strip() if len(row) > 1 else None
    data_origem_str = str(row[2]).strip() if len(row) > 2 else None
    data_destino_str = str(row[3]).strip() if len(row) > 3 else None
    motivo = str(row[4]).strip() if len(row) > 4 else ''
    estado = str(row[5]).strip() if len(row) > 5 else ''
    data_pedido_str = str(row[6]).strip() if len(row) > 6 else None

    b_orig = Bombeiro.query.filter_by(mecanografico=mec_origem).first() if mec_origem else None
    b_dest = Bombeiro.query.filter_by(mecanografico=mec_destino).first() if mec_destino else None
    data_orig = _parse_data(data_origem_str)
    data_dest = _parse_data(data_destino_str)
    data_pedido = _parse_datetime(data_pedido_str) or datetime.utcnow()
    if not b_orig or not b_dest or not data_orig or not data_dest:
        return "Dados de troca incompletos"
    t = TrocaServico(bombeiro_origem_id=b_orig.id, bombeiro_destino_id=b_dest.id,
                     data_origem=data_orig, data_destino=data_dest,
                     motivo=motivo, estado=estado, data_pedido=data_pedido)
    db.session.add(t)
    db.session.flush()
    return None


def _importar_linha_dispensas(row, row_num):
    mec = str(row[0]).strip() if row[0] else None
    inicio_str = str(row[1]).strip() if len(row) > 1 else None
    fim_str = str(row[2]).strip() if len(row) > 2 else None
    motivo = str(row[3]).strip() if len(row) > 3 else ''
    aprovada = str(row[4]).strip().lower() == 'sim' if len(row) > 4 else False
    inicio = _parse_data(inicio_str)
    fim = _parse_data(fim_str)
    if not mec or not inicio or not fim:
        return "Dados de dispensa incompletos"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    d = Dispensa(bombeiro_id=bombeiro.id, data_inicio=inicio, data_fim=fim,
                 motivo=motivo, aprovada=aprovada)
    db.session.add(d)
    db.session.flush()
    return None


def _importar_linha_creditos(row, row_num):
    mec = str(row[0]).strip() if row[0] else None
    data_str = str(row[1]).strip() if len(row) > 1 else None
    descricao = str(row[2]).strip() if len(row) > 2 else ''
    horas = int(row[3]) if len(row) > 3 and row[3] else 8
    estado = str(row[4]).strip() if len(row) > 4 else 'Não Gozado'
    data = _parse_data(data_str)
    if not mec or not data:
        return "Dados de crédito incompletos"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    c = CreditoDispensa(bombeiro_id=bombeiro.id, data=data,
                        descricao=descricao, horas=horas, observacao=estado)
    db.session.add(c)
    db.session.flush()
    return None


def _importar_linha_ecins(row, row_num):
    mec = str(row[0]).strip() if row[0] else None
    data_str = str(row[1]).strip() if len(row) > 1 else None
    turno = str(row[2]).strip() if len(row) > 2 else ''
    categoria = str(row[3]).strip() if len(row) > 3 else ''
    funcao = str(row[4]).strip() if len(row) > 4 else None
    estado = str(row[5]).strip() if len(row) > 5 else 'Pendente'
    data = _parse_data(data_str)
    if not mec or not data:
        return "Dados de ECIN incompletos"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    ec = Ecin(bombeiro_id=bombeiro.id, data=data, turno=turno,
              categoria=categoria, funcao=funcao, estado=estado)
    db.session.add(ec)
    db.session.flush()
    return None


def _importar_linha_fardamentos(row, row_num):
    mec = str(row[1]).strip() if len(row) > 1 and row[1] else None  # coluna 1 = mec
    tipo = str(row[2]).strip() if len(row) > 2 else ''
    nome = str(row[3]).strip() if len(row) > 3 else ''
    tamanho = str(row[4]).strip() if len(row) > 4 else ''
    motivo = str(row[5]).strip() if len(row) > 5 else ''
    estado = str(row[6]).strip() if len(row) > 6 else 'Pedido'
    data_reg_str = str(row[0]).strip() if len(row) > 0 and row[0] else None  # coluna 0 = data registo
    data_reg = _parse_datetime(data_reg_str) or datetime.utcnow()
    if not mec:
        return "Mecanográfico do bombeiro obrigatório"
    bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
    if not bombeiro:
        return f"Bombeiro {mec} não encontrado"
    f = Fardamento(bombeiro_id=bombeiro.id, tipo=tipo, nome=nome,
                   tamanho=tamanho, motivo=motivo, estado=estado, data_registo=data_reg)
    db.session.add(f)
    db.session.flush()
    return None


def _importar_linha_oficina(row, row_num):
    codigo = str(row[0]).strip() if row[0] else ''
    nome_oficina = str(row[1]).strip() if len(row) > 1 else ''
    data_rec_str = str(row[2]).strip() if len(row) > 2 else None
    motivo = str(row[3]).strip() if len(row) > 3 else ''
    avaria_cod = str(row[4]).strip() if len(row) > 4 else None
    matricula = str(row[5]).strip() if len(row) > 5 else None
    kms = int(row[6]) if len(row) > 6 and row[6] else None
    estado = str(row[7]).strip() if len(row) > 7 else 'Oficina'

    if not nome_oficina:
        return "Nome da oficina obrigatório"
    viatura = Viatura.query.filter_by(matricula=matricula).first() if matricula else None
    avaria = Avaria.query.filter_by(codigo=avaria_cod).first() if avaria_cod else None
    data_recepcao = _parse_data(data_rec_str)
    o = Oficina(codigo=codigo, nome_oficina=nome_oficina,
                data_recepcao=data_recepcao,
                motivo=motivo, avaria_id=avaria.id if avaria else None,
                viatura_id=viatura.id if viatura else 1,
                kms=kms, estado=estado)
    db.session.add(o)
    db.session.flush()
    return None


def _importar_linha_gestao_frota(row, row_num):
    matricula = str(row[0]).strip() if row[0] else None
    if not matricula:
        return "Matrícula obrigatória"
    viatura = Viatura.query.filter_by(matricula=matricula).first()
    if not viatura:
        return f"Viatura {matricula} não encontrada"
    g = GestaoFrota(viatura_id=viatura.id)
    g.inspecao_periodica = _parse_data(str(row[1]).strip()) if len(row) > 1 and row[1] else None
    g.kms_ultima_revisao = int(row[2]) if len(row) > 2 and row[2] else None
    g.kms_proxima_revisao = int(row[3]) if len(row) > 3 and row[3] else None
    g.kms_pneus_dianteiros = int(row[4]) if len(row) > 4 and row[4] else None
    g.kms_pneus_trazeiros = int(row[5]) if len(row) > 5 and row[5] else None
    g.kms_correia = int(row[6]) if len(row) > 6 and row[6] else None
    g.outros_apontamentos = str(row[7]).strip() if len(row) > 7 and row[7] else ''
    db.session.add(g)
    db.session.flush()
    return None


def _importar_linha_stock_ambulancia(row, row_num):
    data_str = str(row[0]).strip() if row[0] else None
    matricula = str(row[1]).strip() if len(row) > 1 else None
    nome_produto = str(row[2]).strip() if len(row) > 2 else ''
    quantidade = int(row[3]) if len(row) > 3 and row[3] else 0
    mec_solicitante = str(row[4]).strip() if len(row) > 4 else None
    mec_responsavel = str(row[5]).strip() if len(row) > 5 else None
    confirmado = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False

    viatura = Viatura.query.filter_by(matricula=matricula).first() if matricula else None
    produto = StockFarmacia.query.filter_by(nome=nome_produto).first() if nome_produto else None
    solicitante = Bombeiro.query.filter_by(mecanografico=mec_solicitante).first() if mec_solicitante else None
    responsavel = Bombeiro.query.filter_by(mecanografico=mec_responsavel).first() if mec_responsavel else None

    sa = StockAmbulancia(
        ambulancia_id=viatura.id if viatura else 1,
        produto_id=produto.id if produto else 1,
        quantidade=quantidade,
        solicitante_id=solicitante.id if solicitante else 1,
        responsavel_id=responsavel.id if responsavel else None,
        checklist_id=None,
        confirmado=confirmado,
        data=_parse_datetime(data_str) or datetime.utcnow()
    )
    db.session.add(sa)
    db.session.flush()
    return None


def _importar_linha_notas_central(row, row_num):
    criador_mec = str(row[0]).strip() if row[0] else None
    data_criacao_str = str(row[1]).strip() if len(row) > 1 else None
    descricao = str(row[2]).strip() if len(row) > 2 else ''
    data_evento_str = str(row[3]).strip() if len(row) > 3 else None

    criador = Bombeiro.query.filter_by(mecanografico=criador_mec).first() if criador_mec else None
    data_criacao = _parse_datetime(data_criacao_str) or datetime.utcnow()
    data_evento = _parse_data(data_evento_str)

    n = Nota(criador_id=criador.id if criador else 1, data_criacao=data_criacao,
             descricao=descricao, data_evento=data_evento)
    db.session.add(n)
    db.session.flush()
    return None


def _importar_linha_mensagens_correio(row, row_num):
    remetente_mec = str(row[0]).strip() if row[0] else None
    destinatario_mec = str(row[1]).strip() if len(row) > 1 else None
    departamento = str(row[2]).strip() if len(row) > 2 else None
    assunto = str(row[3]).strip() if len(row) > 3 else ''
    corpo = str(row[4]).strip() if len(row) > 4 else ''
    data_envio_str = str(row[5]).strip() if len(row) > 5 else None
    lida = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False

    remetente = Bombeiro.query.filter_by(mecanografico=remetente_mec).first() if remetente_mec else None
    destinatario = Bombeiro.query.filter_by(mecanografico=destinatario_mec).first() if destinatario_mec else None
    data_envio = _parse_datetime(data_envio_str) or datetime.utcnow()

    m = MensagemCorreio(
        remetente_id=remetente.id if remetente else 1,
        destinatario_id=destinatario.id if destinatario else None,
        departamento=departamento if departamento and departamento != '' else None,
        assunto=assunto, corpo=corpo, data_envio=data_envio, lida=lida
    )
    db.session.add(m)
    db.session.flush()
    return None


@app.route('/backup/importar', methods=['POST'])
@login_required
def backup_importar():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    if 'ficheiro_backup' not in request.files:
        flash('Nenhum ficheiro enviado.', 'warning')
        return redirect(url_for('dashboard'))

    ficheiro = request.files['ficheiro_backup']
    if ficheiro.filename == '' or not ficheiro.filename.endswith(('.xlsx', '.xlsm')):
        flash('Formato inválido. Use .xlsx.', 'danger')
        return redirect(url_for('dashboard'))

    try:
        wb = openpyxl.load_workbook(ficheiro)
    except Exception as e:
        flash(f'Erro ao ler ficheiro: {str(e)}', 'danger')
        return redirect(url_for('dashboard'))

    erros = []
    total_importado = 0

    # ---------- 1. APAGAR TODOS OS DADOS EXISTENTES ----------
    modelos_para_apagar = [
        Nota, MensagemCorreio, ChecklistAmbulanciaItem, ChecklistAmbulancia,
        StockAmbulancia, StockFarmacia, CategoriaFarmacia,
        Fardamento, Ecin, GestaoFrota, Oficina,
        CreditoDispensa, Dispensa, TrocaServico, Escala,
        Avaria, Disponibilidade, Viatura, Bombeiro
    ]
    for modelo in modelos_para_apagar:
        db.session.query(modelo).delete()
    db.session.flush()

    # ---------- 2. IMPORTAR BOMBEIROS ----------
    if 'Bombeiros' in wb.sheetnames:
        ws = wb['Bombeiros']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                numero = str(row[0]).strip() if row[0] else ''
                mecanografico = str(row[1]).strip() if row[1] else ''
                nome = str(row[2]).strip() if row[2] else ''
                nomecompleto = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                email = str(row[4]).strip().lower() if len(row) > 4 and row[4] else ''
                telemovel = str(row[5]).strip() if len(row) > 5 and row[5] else None
                posto = str(row[6]).strip() if len(row) > 6 and row[6] else ''
                tipo_bombeiro = str(row[7]).strip() if len(row) > 7 and row[7] else 'Voluntário'
                departamento = str(row[8]).strip() if len(row) > 8 and row[8] else None
                tipo_user = str(row[9]).strip() if len(row) > 9 and row[9] else 'User'
                ativo = str(row[10]).strip().lower() == 'sim' if len(row) > 10 and row[10] else True
                password_hash = str(row[11]).strip() if len(row) > 11 and row[11] else generate_password_hash('123456')

                if Bombeiro.query.filter(
                    (Bombeiro.numero_interno == numero) |
                    (Bombeiro.mecanografico == mecanografico) |
                    (Bombeiro.email == email)
                ).first():
                    erros.append(f"Bombeiros linha {row_num}: duplicado")
                    continue

                b = Bombeiro(
                    numero_interno=numero,
                    mecanografico=mecanografico,
                    nome=nome,
                    nomecompleto=nomecompleto if nomecompleto else None,
                    email=email,
                    telemovel=telemovel if telemovel and telemovel != '' else None,
                    posto=posto,
                    tipo_bombeiro=tipo_bombeiro,
                    resp_departamento=departamento if departamento and departamento != '' else None,
                    tipo_user=tipo_user,
                    ativo=ativo,
                    password_hash=password_hash
                )
                db.session.add(b)
                total_importado += 1
            except Exception as e:
                erros.append(f"Bombeiros linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 3. IMPORTAR VIATURAS ----------
    if 'Viaturas' in wb.sheetnames:
        ws = wb['Viaturas']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                matricula = str(row[0]).strip() if row[0] else ''
                tipo = str(row[1]).strip() if len(row) > 1 else ''
                nomenclatura = str(row[2]).strip() if len(row) > 2 else ''
                marca = str(row[3]).strip() if len(row) > 3 else ''
                modelo = str(row[4]).strip() if len(row) > 4 else ''
                ano = int(row[5]) if len(row) > 5 and row[5] else 0
                estado = str(row[6]).strip().lower() if len(row) > 6 else 'operacional'
                if Viatura.query.filter_by(matricula=matricula).first():
                    erros.append(f"Viaturas linha {row_num}: matrícula já existe")
                    continue
                v = Viatura(matricula=matricula, tipo=tipo, nomenclatura=nomenclatura,
                            marca=marca, modelo=modelo, ano=ano, estado=estado)
                db.session.add(v)
                total_importado += 1
            except Exception as e:
                erros.append(f"Viaturas linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 4. IMPORTAR CATEGORIAS FARMÁCIA (antes de Stock Farmácia) ----------
    if 'Categorias Farmacia' in wb.sheetnames:
        ws = wb['Categorias Farmacia']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                nome = str(row[0]).strip() if row[0] else ''
                checklist = str(row[1]).strip().lower() == 'sim' if len(row) > 1 else False
                if CategoriaFarmacia.query.filter_by(nome=nome).first():
                    erros.append(f"Categorias Farmácia linha {row_num}: já existe")
                    continue
                c = CategoriaFarmacia(nome=nome, checklist=checklist)
                db.session.add(c)
                total_importado += 1
            except Exception as e:
                erros.append(f"Categorias Farmácia linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 5. IMPORTAR STOCK FARMÁCIA ----------
    if 'Stock Farmacia' in wb.sheetnames:
        ws = wb['Stock Farmacia']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                # Coluna 0: ID original (pode ser None/vazio)
                id_original = int(row[0]) if row[0] is not None else None
                cat = str(row[1]).strip() if len(row) > 1 and row[1] else ''
                nome = str(row[2]).strip() if len(row) > 2 and row[2] else ''
                tamanho_val = str(row[3]).strip()[:100] if len(row) > 3 and row[3] else ''
                stock_val = int(row[4]) if len(row) > 4 and row[4] else 0

                if not nome:
                    erros.append(f"Stock Farmácia linha {row_num}: nome em branco")
                    continue

                # Se for fornecida a data de atualização na coluna 5, usá-la
                data_atualizacao = datetime.utcnow()
                if len(row) > 5 and row[5]:
                    try:
                        data_atualizacao = datetime.strptime(str(row[5]).strip(), '%d/%m/%Y %H:%M')
                    except Exception:
                        pass

                if id_original:
                    # Ver se já existe um registo com este ID
                    existente = StockFarmacia.query.get(id_original)
                    if existente:
                        # Atualiza os campos (mantendo o mesmo ID)
                        existente.categoria = cat
                        existente.nome = nome
                        existente.tamanho = tamanho_val if tamanho_val else None
                        existente.stock = stock_val
                        existente.data_atualizacao = data_atualizacao
                    else:
                        # Cria novo com o ID especificado
                        s = StockFarmacia(
                            id=id_original,
                            categoria=cat,
                            nome=nome,
                            tamanho=tamanho_val if tamanho_val else None,
                            stock=stock_val,
                            data_atualizacao=data_atualizacao
                        )
                        db.session.add(s)
                else:
                    # Sem ID, cria novo com autoincremento
                    s = StockFarmacia(
                        categoria=cat,
                        nome=nome,
                        tamanho=tamanho_val if tamanho_val else None,
                        stock=stock_val,
                        data_atualizacao=data_atualizacao
                    )
                    db.session.add(s)
                total_importado += 1
            except Exception as e:
                erros.append(f"Stock Farmácia linha {row_num}: {str(e)[:150]}")
        db.session.flush()

    # ---------- 6. IMPORTAR STOCK FARDAMENTO ----------
    if 'Stock Fardamento' in wb.sheetnames:
        ws = wb['Stock Fardamento']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                nome = str(row[0]).strip() if row[0] else ''
                descricao = str(row[1]).strip() if len(row) > 1 else ''
                tamanho = str(row[2]).strip() if len(row) > 2 else ''
                tipo = str(row[3]).strip() if len(row) > 3 else ''
                stock = int(row[4]) if len(row) > 4 and row[4] else 0
                s = StockFardamento(nome=nome, descricao=descricao, tamanho=tamanho, tipo=tipo, stock=stock)
                db.session.add(s)
                total_importado += 1
            except Exception as e:
                erros.append(f"Stock Fardamento linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 7. IMPORTAR DISPONIBILIDADES ----------
    if 'Disponibilidades' in wb.sheetnames:
        ws = wb['Disponibilidades']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec = str(row[0]).strip() if row[0] else None
                data_str = str(row[1]).strip() if len(row) > 1 else None
                turno_extra = str(row[2]).strip() if len(row) > 2 else ''
                categoria = str(row[3]).strip() if len(row) > 3 else ''
                confirmada = str(row[4]).strip().lower() == 'sim' if len(row) > 4 else False
                data = _parse_data(data_str) if data_str else None
                if not mec or not data:
                    erros.append(f"Disponibilidades linha {row_num}: dados incompletos")
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"Disponibilidades linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                d = Disponibilidade(bombeiro_id=bombeiro.id, data=data, turno_extra=turno_extra,
                                    categoria=categoria, confirmada=confirmada)
                db.session.add(d)
                total_importado += 1
            except Exception as e:
                erros.append(f"Disponibilidades linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 8. IMPORTAR ESCALAS ----------
    if 'Escalas' in wb.sheetnames:
        ws = wb['Escalas']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec = str(row[0]).strip() if row[0] else None
                inicio_str = str(row[1]).strip() if len(row) > 1 else None
                fim_str = str(row[2]).strip() if len(row) > 2 else None
                turno = str(row[3]).strip() if len(row) > 3 else ''
                categoria = str(row[4]).strip() if len(row) > 4 else 'Bombeiro'
                funcao = str(row[5]).strip() if len(row) > 5 else None
                inicio = _parse_datetime(inicio_str)
                fim = _parse_datetime(fim_str)
                if not mec or not inicio or not fim:
                    erros.append(f"Escalas linha {row_num}: dados incompletos")
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"Escalas linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                e = Escala(bombeiro_id=bombeiro.id, data_inicio=inicio, data_fim=fim,
                           turno=turno, categoria=categoria,
                           funcao=funcao if funcao and funcao != '' else None)
                db.session.add(e)
                total_importado += 1
            except Exception as e:
                erros.append(f"Escalas linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 9. IMPORTAR AVARIAS ----------
    if 'Avarias' in wb.sheetnames:
        ws = wb['Avarias']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                codigo = str(row[0]).strip() if row[0] else None
                matricula = str(row[1]).strip() if len(row) > 1 else None
                descricao = str(row[2]).strip() if len(row) > 2 else ''
                reportador_mec = str(row[3]).strip() if len(row) > 3 else None
                kms = int(row[4]) if len(row) > 4 and row[4] else None
                resp_oficina = str(row[5]).strip().lower() == 'sim' if len(row) > 5 else False
                comando = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False
                estado = str(row[7]).strip() if len(row) > 7 else 'Pendente'
                data_str = str(row[8]).strip() if len(row) > 8 else None
                if not descricao or not matricula:
                    erros.append(f"Avarias linha {row_num}: descrição ou matrícula em falta")
                    continue
                viatura = Viatura.query.filter_by(matricula=matricula).first() if matricula else None
                if not viatura:
                    erros.append(f"Avarias linha {row_num}: viatura {matricula} não encontrada")
                    continue
                reportador = Bombeiro.query.filter_by(mecanografico=reportador_mec).first() if reportador_mec else None
                data_reporte = _parse_datetime(data_str) or datetime.utcnow()
                a = Avaria(codigo=codigo, viatura_id=viatura.id,
                           descricao=descricao, reportado_por=reportador.id if reportador else 1,
                           kms=kms, responsavel_oficina=resp_oficina, comando_verificado=comando,
                           estado=estado, data_reporte=data_reporte)
                db.session.add(a)
                total_importado += 1
            except Exception as e:
                erros.append(f"Avarias linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 10. IMPORTAR TROCAS ----------
    if 'Trocas' in wb.sheetnames:
        ws = wb['Trocas']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec_origem = str(row[0]).strip() if row[0] else None
                mec_destino = str(row[1]).strip() if len(row) > 1 else None
                data_origem_str = str(row[2]).strip() if len(row) > 2 else None
                data_destino_str = str(row[3]).strip() if len(row) > 3 else None
                motivo = str(row[4]).strip() if len(row) > 4 else ''
                estado = str(row[5]).strip() if len(row) > 5 else ''
                data_pedido_str = str(row[6]).strip() if len(row) > 6 else None
                data_orig = _parse_data(data_origem_str)
                data_dest = _parse_data(data_destino_str)
                data_pedido = _parse_datetime(data_pedido_str) or datetime.utcnow()
                if not mec_origem or not mec_destino or not data_orig or not data_dest:
                    erros.append(f"Trocas linha {row_num}: dados incompletos")
                    continue
                b_orig = Bombeiro.query.filter_by(mecanografico=mec_origem).first()
                b_dest = Bombeiro.query.filter_by(mecanografico=mec_destino).first()
                if not b_orig or not b_dest:
                    erros.append(f"Trocas linha {row_num}: bombeiro não encontrado")
                    continue
                t = TrocaServico(bombeiro_origem_id=b_orig.id, bombeiro_destino_id=b_dest.id,
                                 data_origem=data_orig, data_destino=data_dest,
                                 motivo=motivo, estado=estado, data_pedido=data_pedido)
                db.session.add(t)
                total_importado += 1
            except Exception as e:
                erros.append(f"Trocas linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 11. IMPORTAR DISPENSAS ----------
    if 'Dispensas' in wb.sheetnames:
        ws = wb['Dispensas']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec = str(row[0]).strip() if row[0] else None
                inicio_str = str(row[1]).strip() if len(row) > 1 else None
                fim_str = str(row[2]).strip() if len(row) > 2 else None
                motivo = str(row[3]).strip() if len(row) > 3 else ''
                aprovada = str(row[4]).strip().lower() == 'sim' if len(row) > 4 else False
                inicio = _parse_data(inicio_str)
                fim = _parse_data(fim_str)
                if not mec or not inicio or not fim:
                    erros.append(f"Dispensas linha {row_num}: dados incompletos")
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"Dispensas linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                d = Dispensa(bombeiro_id=bombeiro.id, data_inicio=inicio, data_fim=fim,
                             motivo=motivo, aprovada=aprovada)
                db.session.add(d)
                total_importado += 1
            except Exception as e:
                erros.append(f"Dispensas linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 12. IMPORTAR CRÉDITOS ----------
    if 'Créditos' in wb.sheetnames:
        ws = wb['Créditos']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec = str(row[0]).strip() if row[0] else None
                data_str = str(row[1]).strip() if len(row) > 1 else None
                descricao = str(row[2]).strip() if len(row) > 2 else ''
                horas = int(row[3]) if len(row) > 3 and row[3] else 8
                estado = str(row[4]).strip() if len(row) > 4 else 'Não Gozado'
                data = _parse_data(data_str)
                if not mec or not data:
                    erros.append(f"Créditos linha {row_num}: dados incompletos")
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"Créditos linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                c = CreditoDispensa(bombeiro_id=bombeiro.id, data=data,
                                    descricao=descricao, horas=horas, observacao=estado)
                db.session.add(c)
                total_importado += 1
            except Exception as e:
                erros.append(f"Créditos linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 13. IMPORTAR ECINS ----------
    if 'ECINS' in wb.sheetnames:
        ws = wb['ECINS']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                mec = str(row[0]).strip() if row[0] else None
                data_str = str(row[1]).strip() if len(row) > 1 else None
                turno = str(row[2]).strip() if len(row) > 2 else ''
                categoria = str(row[3]).strip() if len(row) > 3 else ''
                funcao = str(row[4]).strip() if len(row) > 4 else None
                estado = str(row[5]).strip() if len(row) > 5 else 'Pendente'
                data = _parse_data(data_str)
                if not mec or not data:
                    erros.append(f"ECINS linha {row_num}: dados incompletos")
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"ECINS linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                ec = Ecin(bombeiro_id=bombeiro.id, data=data, turno=turno,
                          categoria=categoria, funcao=funcao, estado=estado)
                db.session.add(ec)
                total_importado += 1
            except Exception as e:
                erros.append(f"ECINS linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 14. IMPORTAR GESTÃO FROTA ----------
    if 'Gestao Frota' in wb.sheetnames:
        ws = wb['Gestao Frota']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                matricula = str(row[0]).strip() if row[0] else None
                if not matricula:
                    continue
                viatura = Viatura.query.filter_by(matricula=matricula).first()
                if not viatura:
                    erros.append(f"Gestão Frota linha {row_num}: viatura {matricula} não encontrada")
                    continue
                g = GestaoFrota(viatura_id=viatura.id)
                g.inspecao_periodica = _parse_data(str(row[1]).strip()) if len(row) > 1 and row[1] else None
                g.kms_ultima_revisao = int(row[2]) if len(row) > 2 and row[2] else None
                g.kms_proxima_revisao = int(row[3]) if len(row) > 3 and row[3] else None
                g.kms_pneus_dianteiros = int(row[4]) if len(row) > 4 and row[4] else None
                g.kms_pneus_trazeiros = int(row[5]) if len(row) > 5 and row[5] else None
                g.kms_correia = int(row[6]) if len(row) > 6 and row[6] else None
                g.outros_apontamentos = str(row[7]).strip() if len(row) > 7 and row[7] else ''
                db.session.add(g)
                total_importado += 1
            except Exception as e:
                erros.append(f"Gestão Frota linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 15. IMPORTAR FARDAMENTOS ----------
    if 'Fardamentos' in wb.sheetnames:
        ws = wb['Fardamentos']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                data_reg_str = str(row[0]).strip() if row[0] else None
                mec = str(row[1]).strip() if len(row) > 1 else None
                tipo = str(row[2]).strip() if len(row) > 2 else ''
                nome = str(row[3]).strip() if len(row) > 3 else ''
                tamanho = str(row[4]).strip() if len(row) > 4 else ''
                motivo = str(row[5]).strip() if len(row) > 5 else ''
                estado = str(row[6]).strip() if len(row) > 6 else 'Pedido'
                if not mec:
                    continue
                bombeiro = Bombeiro.query.filter_by(mecanografico=mec).first()
                if not bombeiro:
                    erros.append(f"Fardamentos linha {row_num}: bombeiro {mec} não encontrado")
                    continue
                data_reg = _parse_datetime(data_reg_str) or datetime.utcnow()
                f = Fardamento(bombeiro_id=bombeiro.id, tipo=tipo, nome=nome,
                               tamanho=tamanho, motivo=motivo, estado=estado, data_registo=data_reg)
                db.session.add(f)
                total_importado += 1
            except Exception as e:
                erros.append(f"Fardamentos linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 16. IMPORTAR OFICINA ----------
    if 'Oficina' in wb.sheetnames:
        ws = wb['Oficina']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                codigo = str(row[0]).strip() if row[0] else ''
                nome_oficina = str(row[1]).strip() if len(row) > 1 else ''
                data_rec_str = str(row[2]).strip() if len(row) > 2 else None
                motivo = str(row[3]).strip() if len(row) > 3 else ''
                avaria_cod = str(row[4]).strip() if len(row) > 4 else None
                matricula = str(row[5]).strip() if len(row) > 5 else None
                kms = int(row[6]) if len(row) > 6 and row[6] else None
                estado = str(row[7]).strip() if len(row) > 7 else 'Oficina'
                data_recepcao = _parse_data(data_rec_str)
                if not nome_oficina or not data_recepcao:
                    continue
                viatura = Viatura.query.filter_by(matricula=matricula).first() if matricula else None
                avaria = Avaria.query.filter_by(codigo=avaria_cod).first() if avaria_cod else None
                o = Oficina(codigo=codigo, nome_oficina=nome_oficina, data_recepcao=data_recepcao,
                            motivo=motivo, avaria_id=avaria.id if avaria else None,
                            viatura_id=viatura.id if viatura else 1, kms=kms, estado=estado)
                db.session.add(o)
                total_importado += 1
            except Exception as e:
                erros.append(f"Oficina linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 17. IMPORTAR STOCK AMBULÂNCIA ----------
    if 'Stock Ambulância' in wb.sheetnames:
        ws = wb['Stock Ambulância']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                data_str = str(row[0]).strip() if row[0] else None
                matricula = str(row[1]).strip() if len(row) > 1 else None
                nome_produto = str(row[2]).strip() if len(row) > 2 else ''
                quantidade = int(row[3]) if len(row) > 3 and row[3] else 0
                mec_solicitante = str(row[4]).strip() if len(row) > 4 else None
                mec_responsavel = str(row[5]).strip() if len(row) > 5 else None
                confirmado = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False
                if not matricula or not nome_produto:
                    continue
                viatura = Viatura.query.filter_by(matricula=matricula).first() if matricula else None
                produto = StockFarmacia.query.filter_by(nome=nome_produto).first() if nome_produto else None
                solicitante = Bombeiro.query.filter_by(mecanografico=mec_solicitante).first() if mec_solicitante else None
                responsavel = Bombeiro.query.filter_by(mecanografico=mec_responsavel).first() if mec_responsavel else None
                sa = StockAmbulancia(
                    ambulancia_id=viatura.id if viatura else 1,
                    produto_id=produto.id if produto else 1,
                    quantidade=quantidade,
                    solicitante_id=solicitante.id if solicitante else 1,
                    responsavel_id=responsavel.id if responsavel else None,
                    checklist_id=None,
                    confirmado=confirmado,
                    data=_parse_datetime(data_str) or datetime.utcnow()
                )
                db.session.add(sa)
                total_importado += 1
            except Exception as e:
                erros.append(f"Stock Ambulância linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    # ---------- 18. IMPORTAR NOTAS e MENSAGENS (pouco relevantes, mas incluídos) ----------
    if 'Notas Central' in wb.sheetnames:
        ws = wb['Notas Central']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                criador_mec = str(row[0]).strip() if row[0] else None
                data_criacao_str = str(row[1]).strip() if len(row) > 1 else None
                descricao = str(row[2]).strip() if len(row) > 2 else ''
                data_evento_str = str(row[3]).strip() if len(row) > 3 else None
                criador = Bombeiro.query.filter_by(mecanografico=criador_mec).first() if criador_mec else None
                data_criacao = _parse_datetime(data_criacao_str) or datetime.utcnow()
                data_evento = _parse_data(data_evento_str) if data_evento_str else None
                n = Nota(criador_id=criador.id if criador else 1, data_criacao=data_criacao,
                         descricao=descricao, data_evento=data_evento)
                db.session.add(n)
                total_importado += 1
            except Exception as e:
                erros.append(f"Notas Central linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    if 'Mensagens Correio' in wb.sheetnames:
        ws = wb['Mensagens Correio']
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(c is None for c in row):
                continue
            try:
                remetente_mec = str(row[0]).strip() if row[0] else None
                destinatario_mec = str(row[1]).strip() if len(row) > 1 else None
                departamento = str(row[2]).strip() if len(row) > 2 else None
                assunto = str(row[3]).strip() if len(row) > 3 else ''
                corpo = str(row[4]).strip() if len(row) > 4 else ''
                data_envio_str = str(row[5]).strip() if len(row) > 5 else None
                lida = str(row[6]).strip().lower() == 'sim' if len(row) > 6 else False
                remetente = Bombeiro.query.filter_by(mecanografico=remetente_mec).first() if remetente_mec else None
                destinatario = Bombeiro.query.filter_by(mecanografico=destinatario_mec).first() if destinatario_mec else None
                data_envio = _parse_datetime(data_envio_str) or datetime.utcnow()
                m = MensagemCorreio(remetente_id=remetente.id if remetente else 1,
                                    destinatario_id=destinatario.id if destinatario else None,
                                    departamento=departamento if departamento and departamento != '' else None,
                                    assunto=assunto, corpo=corpo, data_envio=data_envio, lida=lida)
                db.session.add(m)
                total_importado += 1
            except Exception as e:
                erros.append(f"Mensagens Correio linha {row_num}: {str(e)[:100]}")
        db.session.flush()

    db.session.commit()

    if erros:
        flash(f'{total_importado} registos importados. {len(erros)} erro(s): ' + '; '.join(erros[:5]), 'warning')
    else:
        flash(f'{total_importado} registos importados com sucesso!', 'success')
    return redirect(url_for('dashboard'))

#-----------------Backup-----------------------

@app.route('/backup/exportar')
@login_required
def backup_exportar():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
    header_font = Font(bold=True)

    # ---- 1. Bombeiros ----
    ws = wb.active
    ws.title = "Bombeiros"
    cabecalhos = ['Nº Interno', 'Mecanográfico', 'Nome', 'Nome Completo', 'Email', 'Telemóvel', 'Posto',
                  'Tipo Bombeiro', 'Resp. Departamento', 'Tipo Utilizador', 'Ativo', 'Password Hash']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for b in Bombeiro.query.order_by(Bombeiro.numero_interno).all():
        ws.append([b.numero_interno, b.mecanografico, b.nome, b.nomecompleto or '', b.email,
                   b.telemovel or '', b.posto, b.tipo_bombeiro, b.resp_departamento or '', b.tipo_user,
                   'Sim' if b.ativo else 'Não', b.password_hash])

    # ---- 2. Viaturas ----
    ws = wb.create_sheet("Viaturas")
    cabecalhos = ['Matrícula', 'Tipo', 'Nomenclatura', 'Marca', 'Modelo', 'Ano', 'Estado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for v in Viatura.query.order_by(Viatura.matricula).all():
        ws.append([v.matricula, v.tipo, v.nomenclatura, v.marca, v.modelo, v.ano, v.estado])

    # ---- 3. Avarias ----
    ws = wb.create_sheet("Avarias")
    cabecalhos = ['Código', 'Viatura Matrícula', 'Descrição', 'Reportado por (mec.)', 'Kms',
                  'Resp. Oficina', 'Comando', 'Estado', 'Data Reporte']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for a in Avaria.query.order_by(Avaria.data_reporte.asc()).all():
        ws.append([a.codigo, a.viatura.matricula if a.viatura else '', a.descricao,
                   a.reportador.mecanografico if a.reportador else '', a.kms or '',
                   'Sim' if a.responsavel_oficina else 'Não', 'Sim' if a.comando_verificado else 'Não',
                   a.estado, a.data_reporte.strftime('%d/%m/%Y %H:%M') if a.data_reporte else ''])

    # ---- 4. Escalas ----
    ws = wb.create_sheet("Escalas")
    cabecalhos = ['Mecanográfico', 'Início', 'Fim', 'Turno', 'Categoria', 'Função']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for e in Escala.query.order_by(Escala.data_inicio.asc()).all():
        ws.append([e.bombeiro.mecanografico if e.bombeiro else '',
                   e.data_inicio.strftime('%d/%m/%Y %H:%M') if e.data_inicio else '',
                   e.data_fim.strftime('%d/%m/%Y %H:%M') if e.data_fim else '',
                   e.turno, e.categoria, e.funcao or ''])

    # ---- 5. Trocas ----
    ws = wb.create_sheet("Trocas")
    cabecalhos = ['Origem (mec.)', 'Destino (mec.)', 'Data Origem', 'Data Destino', 'Motivo', 'Estado', 'Data Pedido']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for t in TrocaServico.query.order_by(TrocaServico.data_pedido.asc()).all():
        ws.append([t.bombeiro_origem.mecanografico if t.bombeiro_origem else '',
                   t.bombeiro_destino.mecanografico if t.bombeiro_destino else '',
                   t.data_origem.strftime('%d/%m/%Y') if t.data_origem else '',
                   t.data_destino.strftime('%d/%m/%Y') if t.data_destino else '',
                   t.motivo or '', t.estado or '',
                   t.data_pedido.strftime('%d/%m/%Y %H:%M') if t.data_pedido else ''])

    # ---- 6. Dispensas ----
    ws = wb.create_sheet("Dispensas")
    cabecalhos = ['Bombeiro (mec.)', 'Início', 'Fim', 'Motivo', 'Aprovada']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for d in Dispensa.query.order_by(Dispensa.data_inicio.asc()).all():
        ws.append([d.bombeiro.mecanografico if d.bombeiro else '',
                   d.data_inicio.strftime('%d/%m/%Y') if d.data_inicio else '',
                   d.data_fim.strftime('%d/%m/%Y') if d.data_fim else '',
                   d.motivo or '', 'Sim' if d.aprovada else 'Não'])

    # ---- 7. Disponibilidades ----
    ws = wb.create_sheet("Disponibilidades")
    cabecalhos = ['Bombeiro (mec.)', 'Data', 'Turno Extra', 'Categoria', 'Confirmada']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for d in Disponibilidade.query.order_by(Disponibilidade.data.asc()).all():
        ws.append([d.bombeiro.mecanografico if d.bombeiro else '',
                   d.data.strftime('%d/%m/%Y') if d.data else '',
                   d.turno_extra or '', d.categoria or '',
                   'Sim' if d.confirmada else 'Não'])

    # ---- 8. Créditos ----
    ws = wb.create_sheet("Créditos")
    cabecalhos = ['Bombeiro (mec.)', 'Data', 'Descrição', 'Horas', 'Estado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for c in CreditoDispensa.query.order_by(CreditoDispensa.data.asc()).all():
        ws.append([c.bombeiro.mecanografico if c.bombeiro else '',
                   c.data.strftime('%d/%m/%Y') if c.data else '',
                   c.descricao or '', c.horas, c.observacao or ''])

    # ---- 9. Stock Fardamento ----
    ws = wb.create_sheet("Stock Fardamento")
    cabecalhos = ['Nome', 'Descrição', 'Tamanho', 'Tipo', 'Stock']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for s in StockFardamento.query.order_by(StockFardamento.nome).all():
        ws.append([s.nome, s.descricao or '', s.tamanho or '', s.tipo, s.stock])

    # ---- 10. Stock Farmácia ----

    ws = wb.create_sheet("Stock Farmacia")
    cabecalhos = ['ID', 'Categoria', 'Nome', 'Tamanho', 'Stock', 'Última Atualização']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos) + 1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for s in StockFarmacia.query.order_by(StockFarmacia.nome).all():
        ws.append([s.id, s.categoria, s.nome, s.tamanho or '', s.stock,
                   s.data_atualizacao.strftime('%d/%m/%Y %H:%M') if s.data_atualizacao else ''])

    # ---- 11. Categorias Farmácia ----
    ws = wb.create_sheet("Categorias Farmacia")
    cabecalhos = ['Nome', 'Checklist']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for cat in CategoriaFarmacia.query.order_by(CategoriaFarmacia.nome).all():
        ws.append([cat.nome, 'Sim' if cat.checklist else 'Não'])

    # ---- 12. Stock Ambulância ----
    ws = wb.create_sheet("Stock Ambulância")
    cabecalhos = ['Data', 'Ambulância', 'Produto', 'Quantidade', 'Solicitante (mec.)', 'Responsável (mec.)', 'Confirmado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for sa in StockAmbulancia.query.order_by(StockAmbulancia.data.asc()).all():
        ws.append([sa.data.strftime('%d/%m/%Y %H:%M') if sa.data else '',
                   sa.ambulancia.matricula if sa.ambulancia else '',
                   sa.produto_stock.nome if sa.produto_stock else '',
                   sa.quantidade,
                   sa.solicitante.mecanografico if sa.solicitante else '',
                   sa.responsavel.mecanografico if sa.responsavel else '',
                   'Sim' if sa.confirmado else 'Não'])

    # ---- 13. Checklist Ambulância ----
    ws = wb.create_sheet("Checklist Ambulância")
    cabecalhos = ['ID Checklist', 'Data/Hora', 'Viatura', 'Bombeiro (mec.)', 'Finalizado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for ch in ChecklistAmbulancia.query.order_by(ChecklistAmbulancia.data_hora.asc()).all():
        ws.append([ch.id, ch.data_hora.strftime('%d/%m/%Y %H:%M') if ch.data_hora else '',
                   ch.viatura.matricula if ch.viatura else '',
                   ch.bombeiro.mecanografico if ch.bombeiro else '',
                   'Sim' if ch.finalizado else 'Não'])

    # ---- 14. Checklist Ambulância Itens ----
    ws = wb.create_sheet("Checklist Itens")
    cabecalhos = ['Checklist ID', 'Produto', 'Quantidade']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for item in ChecklistAmbulanciaItem.query.all():
        ws.append([item.checklist_id, item.produto.nome if item.produto else '', item.quantidade])

    # ---- 15. Mensagens Correio ----
    ws = wb.create_sheet("Mensagens Correio")
    cabecalhos = ['Remetente (mec.)', 'Destinatário (mec.)', 'Departamento', 'Assunto', 'Corpo', 'Data', 'Lida']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for m in MensagemCorreio.query.order_by(MensagemCorreio.data_envio.asc()).all():
        ws.append([m.remetente.mecanografico if m.remetente else '',
                   m.destinatario.mecanografico if m.destinatario else '',
                   m.departamento or '', m.assunto, m.corpo,
                   m.data_envio.strftime('%d/%m/%Y %H:%M') if m.data_envio else '',
                   'Sim' if m.lida else 'Não'])

    # ---- 16. Notas (Central) ----
    ws = wb.create_sheet("Notas Central")
    cabecalhos = ['Criador (mec.)', 'Data Criação', 'Descrição', 'Data Evento']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for n in Nota.query.order_by(Nota.data_criacao.asc()).all():
        ws.append([n.criador.mecanografico if n.criador else '',
                   n.data_criacao.strftime('%d/%m/%Y %H:%M') if n.data_criacao else '',
                   n.descricao, n.data_evento.strftime('%d/%m/%Y') if n.data_evento else ''])

    # ---- 17. ECINS ----
    ws = wb.create_sheet("ECINS")
    cabecalhos = ['Bombeiro (mec.)', 'Data', 'Turno', 'Categoria', 'Função', 'Estado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for ec in Ecin.query.order_by(Ecin.data.asc()).all():
        ws.append([ec.bombeiro.mecanografico if ec.bombeiro else '',
                   ec.data.strftime('%d/%m/%Y') if ec.data else '',
                   ec.turno, ec.categoria or '', ec.funcao or '', ec.estado])

    # ---- 18. Oficina ----
    ws = wb.create_sheet("Oficina")
    cabecalhos = ['Código', 'Nome Oficina', 'Data Recepção', 'Motivo', 'Nº Avaria', 'Viatura', 'Kms', 'Estado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for o in Oficina.query.order_by(Oficina.data_registo.asc()).all():
        ws.append([o.codigo, o.nome_oficina,
                   o.data_recepcao.strftime('%d/%m/%Y') if o.data_recepcao else '',
                   o.motivo or '', o.avaria.codigo if o.avaria else '',
                   o.viatura.matricula if o.viatura else '', o.kms or '', o.estado])

    # ---- 19. Gestão Frota ----
    ws = wb.create_sheet("Gestao Frota")
    cabecalhos = ['Matrícula', 'Inspeção', 'Kms Últ. Revisão', 'Kms Próx. Revisão', 'Kms Pneus Diant.',
                  'Kms Pneus Tras.', 'Kms Correia', 'Apontamentos']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for g in GestaoFrota.query.all():
        v = g.viatura
        ws.append([v.matricula if v else '',
                   g.inspecao_periodica.strftime('%d/%m/%Y') if g.inspecao_periodica else '',
                   g.kms_ultima_revisao or '', g.kms_proxima_revisao or '',
                   g.kms_pneus_dianteiros or '', g.kms_pneus_trazeiros or '',
                   g.kms_correia or '', g.outros_apontamentos or ''])

    # ---- 20. Fardamentos ----
    ws = wb.create_sheet("Fardamentos")
    cabecalhos = ['Data Registo', 'Bombeiro (mec.)', 'Tipo', 'Nome', 'Tamanho', 'Motivo', 'Estado']
    ws.append(cabecalhos)
    for col in range(1, len(cabecalhos)+1):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=1, column=col).font = header_font
    for f in Fardamento.query.order_by(Fardamento.data_registo.asc()).all():
        ws.append([f.data_registo.strftime('%d/%m/%Y %H:%M') if f.data_registo else '',
                   f.bombeiro.mecanografico if f.bombeiro else '',
                   f.tipo, f.nome, f.tamanho, f.motivo, f.estado])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='backup_quartel.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

#----------------------Painel Comando__________________

@app.route('/painel-comando')
@login_required
def painel_comando():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito ao Comando.', 'danger')
        return redirect(url_for('dashboard'))

    hoje = date.today()
    daqui_7_dias = hoje + timedelta(days=7)

    # ----- Total escalados hoje -----
    escalas_hoje = Escala.query.filter(
        func.date(Escala.data_inicio) <= hoje,
        func.date(Escala.data_fim) >= hoje
    ).order_by(Escala.categoria, Escala.turno, Escala.data_inicio).all()

    total_hoje = len(set(e.bombeiro_id for e in escalas_hoje))

    categorias_hoje = {}
    for e in escalas_hoje:
        cat = e.categoria or 'Outros'
        if cat not in categorias_hoje:
            categorias_hoje[cat] = []
        categorias_hoje[cat].append({
            'nome': e.bombeiro.nome,
            'mecanografico': e.bombeiro.mecanografico,
            'turno': e.turno,
            'posto': e.bombeiro.posto
        })

    # ----- Avarias ativas -----
    avarias_ativas = Avaria.query.filter(Avaria.estado.in_(['Pendente', 'Analisar'])).count()

    # ----- Trocas pendentes -----
    trocas_pendentes = TrocaServico.query.filter_by(estado='aceite_colega').count()

    # ----- Dispensas por aprovar -----
    dispensas_pendentes = Dispensa.query.filter_by(aprovada=False).count()

    # ----- Créditos em análise -----
    creditos_analise = CreditoDispensa.query.filter_by(observacao='Em Análise').count()

    # ----- Viaturas Inoperacionais -----
    viaturas_inop = Viatura.query.filter(
        func.lower(Viatura.estado) == 'inoperacional'
    ).order_by(Viatura.matricula).all()

    # ----- Reuniões (próximos 7 dias) -----
    reunioes = Reuniao.query.filter(
        Reuniao.data >= hoje,
        Reuniao.data <= daqui_7_dias
    ).order_by(Reuniao.data.asc(), Reuniao.hora.asc()).all()

    # ----- Notas (data_evento nos próximos 7 dias) -----
    notas = NotaComando.query.filter(
        NotaComando.data_evento >= hoje,
        NotaComando.data_evento <= daqui_7_dias
    ).order_by(NotaComando.data_evento.asc()).all()

    return render_template('painel_comando.html',
                           hoje=hoje,
                           total_hoje=total_hoje,
                           categorias_hoje=categorias_hoje,
                           avarias_ativas=avarias_ativas,
                           trocas_pendentes=trocas_pendentes,
                           dispensas_pendentes=dispensas_pendentes,
                           creditos_analise=creditos_analise,
                           viaturas_inop=viaturas_inop,
                           reunioes=reunioes,
                           notas=notas)

def obter_turno_atual():
    """Retorna o nome do turno (ex: '1 - 00h/08h') correspondente à hora atual."""
    hora = datetime.now().hour
    turnos = [
        ('1 - 00h/08h', (0, 8)),
        ('2 - 08h/16h', (8, 16)),
        ('3 - 16h/24h', (16, 24)),
        ('4 - 11h/19h', (11, 19)),
        ('5 - 10h/18h', (10, 18)),
        ('6 - 07h/19h', (7, 19)),
        ('8 - 08h/20h', (8, 20)),
        ('9 - 20h/08h', (20, 24)),  # 20h-24h parte
        ('7 - 19h/07h', None),       # tratado separadamente
    ]
    # Procurar turnos com intervalo simples
    for nome, (ini, fim) in turnos:
        if nome == '7 - 19h/07h':
            continue
        if ini <= hora < fim:
            return nome
    # Turnos que cruzam a meia-noite
    if hora >= 19 or hora < 7:
        return '7 - 19h/07h'
    if hora >= 20 or hora < 8:
        return '9 - 20h/08h'
    return 'Indefinido'

@app.route('/painel/reuniao/adicionar', methods=['POST'])
@login_required
def adicionar_reuniao():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    data = datetime.strptime(request.form['data'], '%Y-%m-%d').date()
    hora = request.form.get('hora', '')
    assunto = request.form['assunto']
    descricao = request.form.get('descricao', '')
    nova = Reuniao(data=data, hora=hora if hora else None, assunto=assunto, descricao=descricao, criador_id=current_user.id)
    db.session.add(nova)
    db.session.commit()
    flash('Reunião agendada.', 'success')
    return redirect(url_for('painel_comando'))

@app.route('/painel/reuniao/editar/<int:id>', methods=['POST'])
@login_required
def editar_reuniao(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    r = Reuniao.query.get_or_404(id)
    r.data = datetime.strptime(request.form['data'], '%Y-%m-%d').date()
    r.hora = request.form.get('hora', '')
    r.assunto = request.form['assunto']
    r.descricao = request.form.get('descricao', '')
    db.session.commit()
    flash('Reunião atualizada.', 'success')
    return redirect(url_for('painel_comando'))

@app.route('/painel/reuniao/apagar/<int:id>')
@login_required
def apagar_reuniao(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    r = Reuniao.query.get_or_404(id)
    db.session.delete(r)
    db.session.commit()
    flash('Reunião removida.', 'info')
    return redirect(url_for('painel_comando'))

@app.route('/painel/notas/adicionar', methods=['POST'])
@login_required
def adicionar_nota_comando():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    descricao = request.form['descricao']
    data_evento_str = request.form.get('data_evento')
    data_evento = datetime.strptime(data_evento_str, '%Y-%m-%d').date() if data_evento_str else None
    nota = NotaComando(
        criador_id=current_user.id,
        descricao=descricao,
        data_evento=data_evento
    )
    db.session.add(nota)
    db.session.commit()
    flash('Nota do comando registada.', 'success')
    return redirect(url_for('painel_comando'))


@app.route('/painel/notas/editar/<int:id>', methods=['POST'])
@login_required
def editar_nota_comando(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    nota = NotaComando.query.get_or_404(id)
    nota.descricao = request.form['descricao']
    data_evento_str = request.form.get('data_evento')
    nota.data_evento = datetime.strptime(data_evento_str, '%Y-%m-%d').date() if data_evento_str else None
    db.session.commit()
    flash('Nota do comando atualizada.', 'success')
    return redirect(url_for('painel_comando'))


@app.route('/painel/notas/apagar/<int:id>')
@login_required
def apagar_nota_comando(id):
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('painel_comando'))
    nota = NotaComando.query.get_or_404(id)
    db.session.delete(nota)
    db.session.commit()
    flash('Nota do comando removida.', 'info')
    return redirect(url_for('painel_comando'))


@app.route('/admin/apagar-tudo')
@login_required
def apagar_tudo():
    if current_user.tipo_user != 'Admin' and current_user.resp_departamento != 'Comando':
        flash('Acesso restrito.', 'danger')
        return redirect(url_for('dashboard'))

    # Ordem inversa para respeitar as chaves estrangeiras
    modelos_para_apagar = [
        Nota, MensagemCorreio, ChecklistAmbulanciaItem, ChecklistAmbulancia,
        StockAmbulancia, StockFarmacia, CategoriaFarmacia,
        Fardamento, Ecin, GestaoFrota, Oficina,
        CreditoDispensa, Dispensa, TrocaServico, Escala,
        Avaria, Disponibilidade, Viatura, Bombeiro
    ]
    for modelo in modelos_para_apagar:
        db.session.query(modelo).delete()
    db.session.commit()
    flash('Todos os dados foram eliminados com sucesso!', 'success')
    return redirect(url_for('dashboard'))

@app.context_processor
def inject_pendencias():
    pendencias = {}
    if current_user.is_authenticated:
        user = current_user
        total = 0

        if user.tipo_user == 'Admin' or user.resp_departamento == 'Comando':
            pendencias['avarias'] = Avaria.query.filter(Avaria.estado.in_(['Pendente', 'Analisar'])).count()
            pendencias['trocas'] = TrocaServico.query.filter_by(estado='aceite_colega').count()
            pendencias['dispensas'] = Dispensa.query.filter_by(aprovada=False).count()
            pendencias['creditos'] = CreditoDispensa.query.filter_by(observacao='Em Análise').count()
            pendencias['fardamento'] = Fardamento.query.filter_by(estado='Pedido').count()
            pendencias['ecins'] = Ecin.query.filter_by(estado='Pendente').count()
            total = sum(pendencias.values())
        else:
            if user.resp_departamento == 'Oficina':
                pendencias['avarias'] = Avaria.query.filter(Avaria.estado.in_(['Pendente', 'Analisar'])).count()
            if user.resp_departamento == 'Fardamento':
                pendencias['fardamento'] = Fardamento.query.filter_by(estado='Pedido').count()
            if user.resp_departamento == 'Secretaria':
                pendencias['ecins'] = Ecin.query.filter_by(estado='Pendente').count()
            if user.resp_departamento == 'Farmacia':
                pendencias['stock_farmacia'] = StockAmbulancia.query.filter_by(confirmado=False).count()
            if user.resp_departamento == 'Socorrista':
                pendencias['stock_ambulancia'] = StockAmbulancia.query.filter_by(confirmado=False).count()
            total = sum(pendencias.values())

        pendencias['total'] = total
    return dict(pendencias=pendencias)


#if __name__ == '__main__':
    #app.run(debug=True)