# -*- coding: utf-8 -*-
from scrapy import Spider
from scrapy.http import Request

def oportunidad_info(response,value):
    return response.xpath('//strong[text()="' + value + '"]/ancestor::node()[2]/span/text()').extract_first()

class NswSpider(Spider):
    name = 'nsw'
    allowed_domains = ['nsw.gov.au']
    #start_urls =['https://tenders.nsw.gov.au/rms/?event=public.RFT.list']

    def __init__(self,category):
        self.start_urls=[category]

    def parse(self, response):
        oportunidades=response.xpath('//h2/a/@href').extract()
        for oportunidad in oportunidades:
            absolute_url=response.urljoin(oportunidad)
            yield Request(absolute_url, meta={'absolute_url':absolute_url},callback=self.parse_oportunidad)

        next_page_url= response.xpath('//a[text()="Next"]/@href').extract_first()
        absolute_next_page_url=response.urljoin(next_page_url)
        yield Request(absolute_next_page_url)

    def parse_oportunidad(self,response):
        titulo=response.xpath('//h1/text()').extract_first()
        detalles=response.xpath('//*[@id="RFT-Details"]/div/p/text()').extract_first()

        #Informacion Oportunidad
        id=oportunidad_info(response,' RFT ID ')
        tipo=oportunidad_info(response,' RFT Type ')
        fecha_publicacion=oportunidad_info(response,' Published ')
        fecha_limite=oportunidad_info(response,' Closes ')
        categoria=oportunidad_info(response,' Category ')
        agencia=oportunidad_info(response,' Agency ')
        persona_contacto=oportunidad_info(response,' Contact Person ')

        yield{   
            'url':response.meta['absolute_url'],
            'Titulo':titulo,
            'Detalles':detalles,
            'ID':id,
            'Tipo':tipo,
            'Fecha_Publicacion':fecha_publicacion,
            'Fecha_Limite':fecha_publicacion,
            'Categoria':categoria,
            'Agencia':agencia,
            'Persona_Contacto':persona_contacto
        }