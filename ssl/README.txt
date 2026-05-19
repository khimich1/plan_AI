Положите сюда файлы SSL-сертификата (НЕ коммитьте в git):

  fullchain.pem  — цепочка (сертификат + промежуточные CA)
  privkey.pem    — закрытый ключ

Если хостинг выдал .crt и .key — переименуйте или соберите fullchain:
  cat your_domain.crt ca_bundle.crt > fullchain.pem
  cp your_domain.key privkey.pem

Права на сервере:
  chmod 600 privkey.pem
  chmod 644 fullchain.pem
