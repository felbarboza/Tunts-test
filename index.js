const GoogleSpreadSheet = require('google-spreadsheet');
const {promisify} = require('util');

// LINK PARA PLANILHA: https://docs.google.com/spreadsheets/d/1ixOQ2myRoryGn_iWni5t4QEbvMFuv4BFOWfPk49QfBw/edit?usp=sharing

const creds = {
  "type": "service_account",
  "project_id": "tunts-test",
  "private_key_id": "4cc899324f83d7930773a043205fb931e80bf642",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCaIBoJyL2HbimF\n7eU/vE8A/mgpDtgpwq/4P7ssFzb0VJMcEracU6+MR+38JBUUpgWIFcWh7K4MEmAc\ntFxAa93xiKmbuic/BKzM4Wy2CargrhpuVkvtY65ayo/k6jaHwRhCX5M/4rsdggmE\nKEfGd/T5D1BFwxVYJEqMumnFPerFHN9A4hrC8H+WbODFFrKDXJXmZOOpCTbJrvav\n1WQg6pUvkSLcrCLBW9Kh4tNoo7gM7GRcPwxe5454wgDvXb48AyWNc/VYQkZWOdIv\nVlI9JTugjPiLLJ9kruDB2b/hCXY3PGsvEBHBPpq/qK7fm4iq8FvGQFtX+TVDYy6s\nonIKEz6HAgMBAAECggEAO9He2VJRAYoH0sQRWOoBLe0QR3NMAfVe8DboMkY2XaGf\n0WMP/l/awFNAsr7ccb24Yue0Y9MlgGj3Zdy4+4YCSBdXYSpgxixN111dINBiwr7A\nYnfbE2G/j9yT+fDPxmPzQvuufrrFDkBk6ibqKMVxTuObL+B2XdYEG1fU6qnL/8HJ\nUWXF/pZVvvGW1HFQDGRlHdQuVgQeinrilJKhJbobIXaOI3iG48aB8aEfxKJWjoz+\n76YwE/VIsWRAyDzb9yoIJvHurkb0rLnQD57uwOsmqnXmRsoiZQsahQIFgBj3gSqE\nnXNLtjArcjH7RaJnmCtt83EmHo7MZUILO7pky97pKQKBgQDHPQpycsfqiMBA/Se5\nAJQuP2+o6SeLmvSthn2DKuU0cfrZO6l25nq79scj1wGHOOvvV4vgqoEH5CRXxKjo\nf+RJvddTRW3exrW52YRMwTOkX4PcBo5dQ+uGVIAA4rs2QSck3hnMnqC+Lbwi6b+w\nTm5n95ZoJWZtefwJ3+If8oronwKBgQDGCNp/JtDxtmDgrNWAhrA3Hrb9hswha/RX\nSYGwRsoQxOWaaNwztvSyU394GNYDyqbLaeevD+qWb15W86sgOGeCbOKt4HjCfAgc\nPvHBXuEIxJXJmLerM8jblf/xQ2RpDWgmsxwXCTSeO7WH2l3NiVa1wdxxHVvTLIQ6\nxwdHU1UZGQKBgFR0J22EAgIEnZnutVvSRv2jni03R7ABqx2zGJj1IdstRWu3wonI\nANaUMK2cgeVT147IyV4eaDt0FYOutPp428f2VMPTdlMsX/O7pDz02HMgmcA2dzpJ\nhBiY0PmPIlRJIdKa4sy9oN18fXc/JiYR2PLxHCxhTh2xy4hUAoIQSZl5AoGAKM4q\nN4kIBMZPr/vtAk6+gJ0Tl6nu5fQYpOPAlVIA0PPBW8+/j+hjA1uxKE31y1I2jDOG\nScw9ykGobsJGwJzet0E4dBuMxoZIJYnSxsWGGQho1OFi9yP0f0qpMk1wozTgARlm\n8Fg1P2WOQi/8pB1ogIsxoR0rjpfdpz7bgRbqsgECgYAVDejQcfRUn2k7VMI8KZ9T\n6pqx7lmNFts6lgAa8PRPRv1KP/WaDIHCWEI7tiJgcaA/o6GAtm0l3QWE4SyIkWAA\n9J0SSwuK97iFig52ngoA6XGgjIPQWLg+W9l7/cXAo350JUEdFt8ypxT/qzyhEisk\nf6fYWC0nPKw6zc0CkcjnJg==\n-----END PRIVATE KEY-----\n",
  "client_email": "tuntssheetaccess@tunts-test.iam.gserviceaccount.com",
  "client_id": "116293415297204363602",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/tuntssheetaccess%40tunts-test.iam.gserviceaccount.com"
}

async function accessSpreadsheet(){
  const doc = new GoogleSpreadSheet('1ixOQ2myRoryGn_iWni5t4QEbvMFuv4BFOWfPk49QfBw');
  await promisify(doc.useServiceAccountAuth)(creds);
  const info = await promisify(doc.getInfo)();
  const sheet = info.worksheets[0];
  
  const rows = await promisify(sheet.getRows)({
    offset: 1,
  });
  rows.forEach(async (row)=>{
    console.log('Aluno:', row.aluno, ' | Faltas: ', row.faltas, ' | P1, P2, P3: ', row.p1, row.p2, row.p3);
    if(row.faltas>15){
      row.situação='Reprovado por Falta';
      row.notaparaaprovaçãofinal='0';
      console.log('Reprovado por falta');
    }
    else{
      const media = (Number(row.p1) + Number(row.p2) + Number(row.p3))/3;
      console.log('Media: ', media.toFixed(0));
      if(media < 50){
        row.situação='Reprovado por Nota';
        row.notaparaaprovaçãofinal='0';
        console.log('Reprovado por Nota');
      }
      else if(media > 70){
        row.situação='Aprovado';
        row.notaparaaprovaçãofinal='0';
        console.log('Aprovado');
      }
      else {
        var notaaprova = 100-media;
        notaaprova = notaaprova.toFixed(0);
        console.log('Exame Final');
        console.log('Nota para passar: ', notaaprova);
        row.situação='Exame Final';
        row.notaparaaprovaçãofinal=String(notaaprova);
      }
    }
    await row.save();
  });
}

accessSpreadsheet();