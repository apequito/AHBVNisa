import os
from PIL import Image

def gerar_icones(imagem_origem, pasta_destino='static'):
    if not os.path.exists(imagem_origem):
        print(f"Erro: {imagem_origem} não encontrada.")
        return

    img = Image.open(imagem_origem).convert("RGBA")
    # Torna a imagem quadrada (corta ao centro)
    lado = min(img.size)
    img = img.crop(((img.width - lado) // 2, (img.height - lado) // 2, (img.width + lado) // 2, (img.height + lado) // 2))

    tamanhos = {
        'favicon-16x16.png': 16,
        'favicon-32x32.png': 32,
        'apple-touch-icon.png': 180,
        'favicon.ico': [16, 32, 48],  # multi‑tamanho
    }

    for nome, tamanho in tamanhos.items():
        caminho = os.path.join(pasta_destino, nome)
        if nome == 'favicon.ico':
            # Gera ICO com vários tamanhos
            img_ico = img.resize((32, 32), Image.Resampling.LANCZOS)
            img_ico.save(caminho, format='ICO', sizes=[(16,16), (32,32), (48,48)])
        else:
            img_resized = img.resize((tamanho, tamanho), Image.Resampling.LANCZOS)
            img_resized.save(caminho, format='PNG')

    print("Ícones gerados com sucesso!")

if __name__ == "__main__":
    gerar_icones('static/images/BVN.COMP.jpg')