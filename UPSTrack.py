
import requests
import base64
import json
from docxtpl import DocxTemplate

def tranlit_if_neccessary(s):
    letter = True
    for l in s:
        try:
            LetterCode = ord(l.encode('cp1251'))
        except Exception:
            LetterCode = 0
        if 128 <= LetterCode <= 255:
            print(LetterCode)
            return letter
    else:
        return False


class UPS:
    def __init__(self, order_number:int, height:int, width:int, length:int, total_gross_weight:float, exchange_rate:float, freight:int, flag:bool):
        self.order_number = order_number
        self.height = height
        self.width = width
        self.length = length
        self.total_gross_weight = total_gross_weight
        self.exchange_rate = exchange_rate
 #       self.freight = f"{freight:.{2}f}"
        self.freight = freight
        self.flag = flag

    def create(self):
        from googletrans import Translator
        from transliterate import translit

        response = requests.get(
            "https://api.jumpseller.com/v1/orders/status/paid.json?login=ef8cde834d6ae52d628f4655daa78f72&authtoken=f18ce6c94152128475efe4493794f433e8184582a9079d2d4d")
        lst_of_orders = response.json()
        info_dict = {}
        t_usd_value = 0
        t_net_weight = 0

        translator = Translator()
        # print(response.json()[0]['order']['id'])
        order_number_i = -1
        for i in range(len(lst_of_orders)):
            if lst_of_orders[i]['order']['id'] == self.order_number:
                order_number_i = i
                # print(order_number_i)
                break

        # print(client_email, client_name,client_surname)
        for key in lst_of_orders[order_number_i]['order']['shipping_address']:
            info_dict[key] = lst_of_orders[order_number_i]['order']['shipping_address'][key]
        for key in lst_of_orders[order_number_i]['order']['customer']:
            info_dict[key] = lst_of_orders[order_number_i]['order']['customer'][key]
        # info_dict["products"] = client_phone = lst_of_orders[order_number_i]['order']['products']
        info_dict["products"] = lst_of_orders[order_number_i]['order']['products']
        info_dict2 = {'height': self.height, 'width': self.width, 'length': self.length, 'total_gross_weight': self.total_gross_weight,
                      'freight': self.freight, 'order_number': self.order_number}
        # добавление словаря info_dict2 в info_dict
        for k, v in info_dict2.items():
            info_dict[k] = v

        info_dict["table_contents"] = lst_of_orders[order_number_i]['order']['products']
        # обработка значений таблицы

        for i in range(len(info_dict["table_contents"])):
            info_dict["table_contents"][i]['usd_value'] = round(
                (info_dict["table_contents"][i]['price'] / self.exchange_rate), 2)

            info_dict["table_contents"][i]['s_usd_value'] = round(
                (info_dict["table_contents"][i]['usd_value'] * info_dict["table_contents"][i]['qty']), 2)

              # проверяем написанно ли имя на кирилице и транслитерируем
            if tranlit_if_neccessary(info_dict["table_contents"][i]['name']):
                info_dict["table_contents"][i]['name'] = translator.translate(
                    info_dict["table_contents"][i]['name']).text

            t_usd_value += info_dict["table_contents"][i]['s_usd_value']

            info_dict["table_contents"][i]['weight'] = round((info_dict["table_contents"][i]['weight'] / 1.24), 1)
            t_net_weight += info_dict["table_contents"][i]['weight'] * info_dict["table_contents"][i]['qty']
            #    info_dict["table_contents"][i]['s_usd_value'] = f"{2.0:.{3}f}"
            w = info_dict["table_contents"][i]['weight']
            info_dict["table_contents"][i]['weight_000'] = f"{w:.{3}f}"
            #   w2 = w * info_dict['table_contents'][i]['qty']
            info_dict["table_contents"][i]['s_weight_000'] = f"{w * info_dict['table_contents'][i]['qty']:.{3}f}"

        # print(info_dict["table_contents"][i])
        info_dict['t_usd_value'] = round(t_usd_value + float(self.freight), 2)
        info_dict['t_net_weight'] = t_net_weight
        info_dict['order_number'] = self.order_number
        # translation if necessary

        if tranlit_if_neccessary(info_dict['name']):
            info_dict['name'] = translit(info_dict['name'], "ru", reversed=True).upper()
            info_dict['surname'] = translit(info_dict['surname'], "ru", reversed=True).upper()
        else:
            info_dict['name'] = info_dict['name'].upper()
            info_dict['surname'] = info_dict['surname'].upper()
            # print(info_dict['name'].upper(), info_dict['surname'].upper() )

        if tranlit_if_neccessary(info_dict['address']):
            info_dict['address'] = translit(info_dict['address'], "ru", reversed=True)[:35]

        if tranlit_if_neccessary(info_dict['city']):
             info_dict['city'] = translator.translate(info_dict['city']).text

        if tranlit_if_neccessary(info_dict['region']):

            info_dict['region'] = translator.translate(info_dict['region']).text

        if tranlit_if_neccessary(info_dict['country']):
            info_dict['country'] = translator.translate(info_dict['country']).text.upper()

        # добавляем день создания заказа
        day = lst_of_orders[order_number_i]['order']['completed_at'][8:10]
        month = lst_of_orders[order_number_i]['order']['completed_at'][5:7]
        year = lst_of_orders[order_number_i]['order']['completed_at'][0:4]
        info_dict["day"] = f'{day}/{month}/{year}'
        headers = {
            'AccessLicenseNumber': '2DB76C505E70C260',
            'Password': 'Sucesso@2022',
            'Content-Type': 'application/json',
            'transId': 'Transaction123',
            'transactionSrc': 'GG',
            'Username': 'canopeiabr',
            'Accept': 'application/json'
        }
        body = {
            "ShipmentRequest": {
                "Shipment": {
                    "Description": "Hair Cosmetics",
                    "Shipper": {
                        "Name": "Canopeia",
                        "AttentionName": "Serge Demidov",
                        "TaxIdentificationNumber": "34554968000181",
                        "Phone": {
                            "Number": "5511964427723"
                        },
                        "ShipperNumber": "723F8F",
                        "Address": {
                            "AddressLine": "R. J. Martins Fernandes 601 Gal 18",
                            "City": "Sao Bernardo",
                            "StateProvinceCode": "SP",
                            "PostalCode": "09843400",
                            "CountryCode": "BR"
                        }
                    },
                    "ShipTo": {
                        "Name": f"{info_dict['name']} {info_dict['surname']}",
                        "AttentionName": f"{info_dict['name']} {info_dict['surname']}",
                        "Phone": {
                            "Number": str(info_dict['phone'])
                        },
                        "FaxNumber": str(info_dict['phone']),
                        "TaxIdentificationNumber": "",
                        "Address": {
                            "AddressLine": info_dict['address'],
                            "City": info_dict['city'],
                            "StateProvinceCode": info_dict['region'][0:2],
                            "PostalCode": info_dict['postal'],
                            "CountryCode": info_dict['country_code']
                        }
                    },
                    "ShipFrom": {
                        "Name": "Canopeia",
                        "AttentionName": "Serge Demidov",
                        "Phone": {
                            "Number": "5511964427723"
                        },
                        "FaxNumber": "5511964427723",
                        "TaxIdentificationNumber": "34554968000181",
                        "Address": {
                            "AddressLine": "R. J. Martins Fernandes 601 Gal 18",
                            "City": "Sao Bernardo",
                            "StateProvinceCode": "SP",
                            "PostalCode": "09843400",
                            "CountryCode": "BR"
                        }
                    },
                    "PaymentInformation": {
                        "ShipmentCharge": {
                            "Type": "01",
                            "BillShipper": {
                                "AccountNumber": "723F8F"
                            }
                        }
                    },
                    "Service": {
                        "Code": "65",
                        "Description": "UPS Saver"
                    },
                    "Package": [
                        {
                            "Description": "Hair Cosmetics",
                            "Packaging": {
                                "Code": "02"
                            },
                            "PackageWeight": {
                                "UnitOfMeasurement": {
                                    "Code": "KGS"
                                },
                                "Weight": str(self.total_gross_weight)
                            },
                            "PackageServiceOptions": ""
                        }
                    ],
                    "ItemizedChargesRequestedIndicator": "",
                    "RatingMethodRequestedIndicator": "",
                    "TaxInformationIndicator": "",
                    "ShipmentRatingOptions": {
                        "NegotiatedRatesIndicator": ""
                    }
                },
                "LabelSpecification": {
                    "LabelImageFormat": {
                        "Code": "GIF"
                    }
                }
            }
        }
        if self.flag:
        # Production URL
            url = 'https://onlinetools.ups.com/ship/v1/shipments?'
        else:
        # Testing URL
            url = "https://wwwcie.ups.com/ship/v1/shipments"

        response = requests.post(url, headers=headers, data=json.dumps(body))
        print(response.status_code)

        info_dict['tracking_number'] = response.json()['ShipmentResponse']['ShipmentResults']['PackageResults'][
            'TrackingNumber']

        # print(response.json())

        image_result = open(f"uploads/{self.order_number}_lable.gif", 'wb')  # create a writable image and write the decoding result
        image_result.write(base64.b64decode(bytes(
            response.json()['ShipmentResponse']['ShipmentResults']['PackageResults']['ShippingLabel']['GraphicImage'],'utf-8')))
        with open("info_dict.json", "w", encoding='utf-8') as outfile:
            json.dump(info_dict, outfile)
        print(response.json()['ShipmentResponse']['ShipmentResults']['ShipmentCharges']['TransportationCharges'][
                  'MonetaryValue'], " USD стбез скидки")
        print(response.json()['ShipmentResponse']['ShipmentResults']['NegotiatedRateCharges']['TotalCharge'][
                  'MonetaryValue'], " USD со скидкой")

        invoice = DocxTemplate("invoice_template.docx")
        pl = DocxTemplate("PL_template.docx")
        invoice.render(info_dict)
        pl.render(info_dict)
        invoice.save(f"uploads/{info_dict['order_number']}_invoice.docx")
        filename = f"{info_dict['order_number']}_invoice.docx"
        pl.save(f"uploads/{info_dict['order_number']}_PL.docx")
        self.track = info_dict['tracking_number']

if __name__ == '__main__':
    UPS.run(debug=False)
