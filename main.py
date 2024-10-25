from json import loads
from os import remove
from os.path import exists

from pandas import DataFrame
from requests import post


class Parser:
    def __init__(self):
        self.headers = {
            'Accept': 'application/json, text/plain, * / *',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'Content-Length': '450',
            'Content-Type': 'application/json',
            'Origin': 'https://online.metro-cc.ru',
            'Priority': 'u=1, i',
            'Referer': 'https://online.metro-cc.ru/',
            'Sec-Ch-Ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Opera GX";v="114"',
            'Sec-Ch-Ua-Mobile': '?0',
            'Sec-Ch-Ua-Platform': '"Windows"',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-site',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 OPR/114.0.0.0',
        }
        self.cookies = {
            '_slid_server': '671a3e4f5f3876392c0679c8',
            'pdp_abc_20': '0',
            'plp_bmpl_bage': '1',
            '_slid': '671a3e4f5f3876392c0679c8',
            '_slfreq': '633ff97b9a3f3b9e90027740%3A633ffa4c90db8d5cf00d7810%3A1729780336%3B64a81e68255733f276099da5%3A64abaf645c1afe216b0a0d38%3A1729780336',
            'metro_user_id': 'f738b959b920d4897f72b1227d1e48e2',
            '_ym_uid': '1729773136795667698',
            '_ym_d': '1729773136',
            '_ym_isad': '1',
            'uxs_uid': '00a62e30-9204-11ef-8a20-bda8a1f8fa6b',
            'allowedCookieCategories': 'necessary%7Cfunctional%7Cperformance%7Cpromotional%7Cthirdparty%7CUncategorized',
            '_gcl_au': '1.1.1755588705.1729779939',
            '_gid': 'GA1.2.361303038.1729779940',
            'local_ga': 'GA1.1.1691554566.1729779940',
            'local_ga_LV657197KB': 'GS1.1.1729779939.1.1.1729780031.34.0.0',
            '_ga': 'GA1.1.1691554566.1729779940',
            'metro_api_session': 'fNpCebxpxzJqs4guHZ7HyPA4SMiu9pEn5pwUFKiD',
            '_ym_visorc': 'b',
            'metroStoreId': '10',
            '_slsession': '91CFF826-0683-45D5-AA2E-76A4FECD6B46',
            '_ga_VHKD93V3FV': 'GS1.1.1729786381.4.1.1729788969.0.0.0',
        }
        self.cities = {
            'Москва': {
                'storeId': 10,
                'size': 415,
            },
            'Санкт-Петербург': {
                'storeId': 15,
                'size': 401,
            },
        }
        self.url = 'https://api.metro-cc.ru/products-api/graph'
        self.file_name = 'Результат.xlsx'
        self.output = []

    def parse(self):
        for city, city_info in self.cities.items():
            data = f'{{"query": "\\n query Query($storeId: Int!, $slug: String!, $attributes:[AttributeFilter], $filters: [FieldFilter], $from: Int!, $size: Int!, $sort: InCategorySort, $in_stock: Boolean, $eshop_order: Boolean, $is_action: Boolean, $priceLevelsOnline: Boolean) {{\\n category (storeId: $storeId, slug: $slug, inStock: $in_stock, eshopAvailability: $eshop_order, isPromo: $is_action, priceLevelsOnline: $priceLevelsOnline) {{\\n id\\n name\\n slug\\n id\\n parent_id\\n meta {{\\n description\\n h1\\n title\\n keywords\\n }}\\n disclaimer\\n description {{\\n top\\n main\\n bottom\\n }}\\n breadcrumbs {{\\n category_type\\n id\\n name\\n parent_id\\n parent_slug\\n slug\\n }}\\n promo_banners {{\\n id\\n image\\n name\\n category_ids\\n type\\n sort_order\\n url\\n is_target_blank\\n analytics {{\\n name\\n category\\n brand\\n type\\n start_date\\n end_date\\n }}\\n }}\\n\\n\\n dynamic_categories(from: 0, size: 9999) {{\\n slug\\n name\\n id\\n category_type\\n dynamic_product_settings {{\\n attribute_id\\n max_value\\n min_value\\n slugs\\n type\\n }}\\n }}\\n filters {{\\n facets {{\\n key\\n total\\n filter {{\\n id\\n hru_filter_slug\\n is_hru_filter\\n is_filter\\n name\\n display_title\\n is_list\\n is_main\\n text_filter\\n is_range\\n category_id\\n category_name\\n values {{\\n slug\\n text\\n total\\n }}\\n }}\\n }}\\n }}\\n total\\n prices {{\\n max\\n min\\n }}\\n pricesFiltered {{\\n max\\n min\\n }}\\n products(attributeFilters: $attributes, from: $from, size: $size, sort: $sort, fieldFilters: $filters)  {{\\n health_warning\\n limited_sale_qty\\n id\\n slug\\n name\\n name_highlight\\n article\\n new_status\\n main_article\\n main_article_slug\\n is_target\\n category_id\\n category {{\\n name\\n }}\\n url\\n images\\n pick_up\\n rating\\n icons {{\\n id\\n badge_bg_colors\\n rkn_icon\\n caption\\n type\\n is_only_for_sales\\n caption_settings {{\\n colors\\n text\\n }}\\n sort\\n image_svg\\n description\\n end_date\\n start_date\\n status\\n }}\\n manufacturer {{\\n name\\n }}\\n packing {{\\n size\\n type\\n }}\\n stocks {{\\n value\\n text\\n scale\\n eshop_availability\\n prices_per_unit {{\\n old_price\\n offline {{\\n price\\n old_price\\n type\\n offline_discount\\n offline_promo\\n }}\\n price\\n is_promo\\n levels {{\\n count\\n price\\n }}\\n online_levels {{\\n count\\n  price\\n discount\\n }}\\n discount\\n }}\\n prices {{\\n price\\n is_promo\\n old_price\\n offline {{\\n old_price\\n price\\n type\\n offline_discount\\n offline_promo\\n }}\\n levels {{\\n count\\n price\\n }}\\n online_levels {{\\n count\\n price\\n discount\\n }}\\n discount\\n }}\\n }}\\n }}\\n argumentFilters {{\\n eshopAvailability\\n inStock\\n isPromo\\n priceLevelsOnline\\n }}\\n }}\\n }}\\n","variables":{{"storeId":{city_info['storeId']},"sort":"default","size":{city_info['size']},"from":0,"filters":[{{"field":"main_article","value":"0"}}],"attributes":[],"in_stock":false,"eshop_order":false,"allStocks":false,"slug":"pitevaya"}}}}'
            response = post(self.url, headers=self.headers, cookies=self.cookies, data=data)
            response = loads(response.text)

            for product in response['data']['category']['products']:
                # Правильнее, мне кажется, сравнивать 'value' с 0, но я точно не уверен в значении этого атрибута,
                # поэтому сравниваю 'text'
                if product['stocks'][0]['text'] == 'Отсутствует':
                    continue

                id = product['article']
                name = product['name']
                link = f'https://online.metro-cc.ru{product['url']}'
                brand = product['manufacturer']['name']
                price = product['stocks'][0]['prices_per_unit']

                if price['is_promo']:
                    regular_price = price['old_price']
                    promo_price = price['price']
                else:
                    regular_price = price['price']
                    promo_price = None

                product_data = [city, id, name, link, regular_price, promo_price, brand]
                self.output.append(product_data)

    def save_xlsx(self):
        header = [
            'Город',
            'ID',
            'Наименование',
            'Ссылка',
            'Регулярная цена',
            'Промо цена',
            'Бренд',
        ]

        df = DataFrame(self.output, columns=header)

        if exists(self.file_name):
            remove(self.file_name)

        df.to_excel(self.file_name, sheet_name='Вода', index=False)


def main():
    parser = Parser()
    parser.parse()
    parser.save_xlsx()


if __name__ == '__main__':
    main()
