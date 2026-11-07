import unittest
from unittest.mock import Mock

from app import GraphClient, validate_config, validate_record

CFG = dict(tenant_id='11111111-1111-1111-1111-111111111111',
           client_id='22222222-2222-2222-2222-222222222222',
           site_url='https://empresa.sharepoint.com/sites/Equipe', list_name='TESTEB2B')


def response(value, status=200):
    obj = Mock(ok=200 <= status < 300, status_code=status, content=b'json')
    obj.json.return_value = value
    return obj


class Tests(unittest.TestCase):
    def client(self):
        client = GraphClient(CFG, lambda: 'test-token')
        client.session = Mock()
        client.base = '/sites/site/lists/list'
        client.columns = {'IDCLIENTE': 'ClienteInterno', 'NOME': 'Title', 'DATA': 'DataInterna'}
        return client

    def test_configuration_rejects_personal_and_list_url(self):
        for url in ['https://lists.live.com/?listId=1', 'http://empresa.sharepoint.com',
                    'https://empresa.sharepoint.com.evil.test/sites/E',
                    CFG['site_url'] + '/Lists/TESTEB2B/AllItems.aspx']:
            with self.assertRaises(ValueError):
                validate_config({**CFG, 'site_url': url})
        self.assertEqual(validate_config(CFG), CFG)

    def test_validation_preserves_leading_zero_and_rejects_bad_date(self):
        self.assertEqual(validate_record(dict(IDCLIENTE='0001', NOME='Ana', DATA='2026-11-07')),
                         dict(IDCLIENTE='0001', NOME='Ana', DATA='2026-11-07T00:00:00Z'))
        with self.assertRaises(ValueError):
            validate_record(dict(IDCLIENTE='1', NOME='Ana', DATA='2026-02-30'))

    def test_pagination_and_rows(self):
        c = self.client()
        c.session.request.side_effect = [response({'value': [dict(id='3', eTag='v1', fields={
            'ClienteInterno': '0001', 'Title': '<b>Ana</b>', 'DataInterna': '2026-11-07T00:00:00Z'})],
            '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next'}), response({'value': []})]
        self.assertEqual(c.rows(), [dict(id='3', etag='v1', IDCLIENTE='0001', NOME='<b>Ana</b>', DATA='2026-11-07')])
        self.assertEqual(c.session.request.call_count, 2)

    def test_create_update_delete_and_conflict(self):
        c = self.client()
        c.session.request.return_value = response({'id': '3'})
        record = dict(IDCLIENTE='0001', NOME='Ana', DATA='2026-11-07')
        c.save(record)
        kwargs = c.session.request.call_args.kwargs
        self.assertEqual(kwargs['json']['fields']['ClienteInterno'], '0001')
        c.save(record, '3', 'v1')
        self.assertEqual(c.session.request.call_args.args[0], 'PATCH')
        self.assertEqual(c.session.request.call_args.kwargs['headers']['If-Match'], 'v1')
        c.session.request.return_value = response({}, 412)
        with self.assertRaisesRegex(ValueError, 'outra pessoa'):
            c.delete('3', 'v1')
        self.assertEqual(c.session.request.call_args.args[0], 'DELETE')

    def test_discovery_maps_display_names_to_internal_names(self):
        c = self.client()
        c.session.request.side_effect = [response({'id': 'site'}),
            response({'value': [{'id': 'list', 'displayName': 'TESTEB2B'}]}),
            response({'value': [dict(displayName='IDCLIENTE', name='ClienteInterno', text={}),
                                dict(displayName='NOME', name='Title', text={}),
                                dict(displayName='DATA', name='DataInterna', dateTime={'format': 'dateOnly'})]})]
        c.connect()
        self.assertEqual(c.columns['NOME'], 'Title')

    def test_pagination_rejects_foreign_host(self):
        c = self.client()
        with self.assertRaises(ValueError):
            c.request('GET', 'https://evil.test/token')
        c.session.request.assert_not_called()


if __name__ == '__main__':
    unittest.main()
