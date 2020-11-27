import strings

class Producto_GPG:
    # Atributos básicos, determinados mediante el formato de origen
    nombre = ''
    descripcion =''
    descripcion_corta = ''
    precio_normal = ''
    categorias = ''
    familia = ''
    mayoreo_a_partir = 1
    precio_mayoreo = 0.0
    existencia = 1
    marca = ''
    atributo = ''
    cantidad = 0
    unidad = ''
    url_img = ''

    # Campos a llenar en el csv de carga masiva
    id = 0          # Identificador único
    tipo = ''       # variable, variation o simple
    sku = ''        # sku para efectos de la plataforma
    #nombre
    publicado = '1'
    destacado = '0'
    visibilidad_catalogo = 'visible'
    #desc_corta
    #desc
    dia_inicio_rebaja = ''
    dia_fin_rebaja = ''
    estado_impuesto = 'taxable'
    clase_impuesto = ''
    en_inventario = '1';
    inventario = ''
    cantidad_bajo_inventario = ''
    permitir_reserva_producto_agotado = ''
    vendido_individualmente = '0'
    peso_kg = ''
    longitud_cm = ''
    anchura_cm = ''
    altura_cm = ''
    permitir_valoraciones_clientes = '1'
    nota_de_compra = ''
    precio_rebajado = ''
    #categorias
    etiquetas = ''
    clase_envio = ''
    imagenes = ''
    limite_descargas = ''
    dias_caducidad = ''
    superior = 0    # valor de familia en csv de plataforma
    productos_agrupados = ''
    ventas_dirigidas = ''
    ventas_cruzadas = ''
    url_externa = ''
    texto_boton = ''
    posicion = 0    # posicion del artículo en la familia
    fixed_tiered_prices = ''
    nombre_atributo_1 = ''
    valor_atributo_1 = ''
    atributo_visible_1 = ''
    atributo_global_1 = ''
    meta_wcz_pps_price_prefix = ''
    meta_site_sidebar_layout = ''
    meta_site_content_layout = ''
    meta_theme_transparent_header_meta = ''
    meta_precio_menudeo = ''
    meta__precio_menudeo = strings.meta__precio_menudeo_str
    meta_precio_mayoreo = 'No aplica'
    meta__precio_mayoreo = strings.meta__precio_mayoreo_str
    nombre_atributo_2 = ''
    valor_atributo_2 = ''
    atributo_visible_2 = ''
    atributo_global_2 = ''
    meta_wp_page_template = ''
    nombre_atributo_3 = ''
    valor_atributo_3 = ''
    atributo_visible_3 = ''
    atributo_global_3 = ''

    def __init__(self,id,sku,nom,desc,desc_c,pu,cat,fam,may,pm,exis,mar,atr,cnt,uni,img):
        self.id = id
        self.sku = sku                      # sku para efectos de la plataforma
        self.nombre = nom
        self.descripcion = desc
        self.descripcion_corta = desc_c
        self.precio_normal = pu
        self.categorias = cat
        self.familia = fam 
        self.mayoreo_a_partir = may
        self.precio_mayoreo = pm
        self.existencia = exis
        self.marca = mar
        self.atributo = atr
        self.cantidad = cnt
        self.unidad = uni
        self.url_img = img

        self.tipo = ''       # variable, variation o simple
        #nombre
        self.publicado = '1'
        self.destacado = '0'
        self.visibilidad_catalogo = 'visible'
        #desc_corta
        #desc
        self.dia_inicio_rebaja = ''
        self.dia_fin_rebaja = ''
        self.estado_impuesto = 'taxable'
        self.clase_impuesto = ''
        self.en_inventario = str(exis);
        self.inventario = ''
        self.cantidad_bajo_inventario = ''
        self.permitir_reserva_producto_agotado = '0'
        self.vendido_individualmente = '0'
        self.peso_kg = ''
        self.longitud_cm = ''
        self.anchura_cm = ''
        self.altura_cm = ''
        self.permitir_valoraciones_clientes = '1'
        self.nota_de_compra = ''
        self.precio_rebajado = ''
        #precio_normal
        #categorias
        self.etiquetas = ''
        self.clase_envio = ''
        self.imagenes = ''
        self.limite_descargas = ''
        self.dias_caducidad = ''
        self.superior = ''    # valor de familia en csv de plataforma
        self.productos_agrupados = ''
        self.ventas_dirigidas = ''
        self.ventas_cruzadas = ''
        self.url_externa = ''
        self.texto_boton = ''
        self.posicion = 0    # posicion del artículo en la familia
        self.fixed_tiered_prices = ''
        self.nombre_atributo_1 = 'PRESENTACION'
        self.valor_atributo_1 = ''
        self.atributo_visible_1 = ''
        self.atributo_global_1 = '0'
        self.meta_wcz_pps_price_prefix = ''
        self.meta_site_sidebar_layout = 'default'
        self.meta_site_content_layout = 'default'
        self.meta_theme_transparent_header_meta = 'default'
        self.meta_precio_menudeo = 'No aplica'
        self.meta__precio_menudeo = strings.meta__precio_menudeo_str
        self.meta_precio_mayoreo = 'No aplica'
        self.meta__precio_mayoreo = strings.meta__precio_mayoreo_str
        self.nombre_atributo_2 = 'ATRIBUTO'
        self.valor_atributo_2 = ''
        self.atributo_visible_2 = ''
        self.atributo_global_2 = '0'
        self.meta_wp_page_template = ''
        self.nombre_atributo_3 = ''
        self.valor_atributo_3 = ''
        self.atributo_visible_3 = ''
        self.atributo_global_3 = ''

    def print_prod(self):
        attrs = vars(self)
        print(' '.join('%s-> %s' %item for item in attrs.items()))


    # id = 0
    # sku = ''
    # clasificacion = ''
    # superior = ''
    # cantidad = 0
    # unidad = ''
    # marca = ''
    # atributo = ''
    # imagen = ''

    # Nuevos parámetros a determinar
    # sku_new = ''
    # superior_new = ''
    # tipo = ''
    # posicion = '-'

    # def __init__(self,id,sku,clasif,sup,cnt,uni,marc,atr,im):
    #     self.id = id
    #     self.sku = sku
    #     self.clasificacion = clasif
    #     self.superior = sup
    #     self.cantidad = cnt
    #     self.unidad = uni
    #     self.marca = marc
    #     self.atributo = atr
    #     self.imagen = im
#End class Producto_GPG