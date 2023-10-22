const Imap = require('imap');
const cheerio = require('cheerio');
const nodemailer = require('nodemailer');
const quotedPrintable = require('quoted-printable');
const WooCommerceAPI = require('woocommerce-api');
const cron = require('node-cron');

const WooCommerce = new WooCommerceAPI({
  url: 'https://kunstinjekeuken.nl/',
  consumerKey: 'ck_6628c66a1c3f7b088cff069580885a133195b3cb',
  consumerSecret: 'cs_74f513b6a491e2466bc906c10faa749d093f7954',
  wpAPI: true,
  version: 'wc/v3'
});

let imap = new Imap({
    user: 'leila.haqi@gmail.com',
    password: 'diid ezch goue hayw',
  host: 'imap.gmail.com',
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false },
  connTimeout: 10000
});

function openInbox(cb) {
  imap.openBox('INBOX', false, cb);
}

function checkEmails() {
  imap.once('ready', function() {
    openInbox(function(err, box) {
      if (err) {
        console.error('Error opening inbox:', err);
        return;
      }
      imap.search(['UNSEEN', ['FROM', 'info@drukland.nl']], function(err, results) {
        if (err) throw err;
        if (results.length === 0) {
          console.log('No new emails to fetch');
          return; // Return from the function without doing anything
        }
        let f = imap.fetch(results, { bodies: '', markSeen: true });
        f.on('message', function(msg, seqno) {
          let prefix = '(#' + seqno + ') ';
          msg.on('body', function(stream, info) {
            let buffer = '';
            stream.on('data', function(chunk) {
              buffer += chunk.toString('utf8');
            });
            stream.once('end', function() {
              let cleanedBuffer = buffer.replace(/=\r\n/g, '');
              let decodedBuffer = quotedPrintable.decode(cleanedBuffer, 'utf8');
              let $ = cheerio.load(decodedBuffer);
              let links = [];
              $('a').each((i, link) => {
                let href = $(link).attr('href');
                if (href && href.includes('https://www.dhlparcel.nl/nl/zakelijk/zending')) {
                  let aTag = `<a href="${href}">${href}</a>`;
                  links.push(aTag);
                }
              });
              console.log('Links:', links);

              // Check if any tracking links were found
              if (links.length === 0) {
                console.log('No tracking links found in email. Not sending email.');
                return;
              }

              let referenceNumberElement = $('p:contains("Referentie")').find('strong').first();
              let referenceNumber = referenceNumberElement.text().replace('#', '');
              console.log('Reference number: ' + referenceNumber);
              referenceNumber = referenceNumber.replace(/\D+$/, '');

              WooCommerce.get(`orders/${referenceNumber}`, function(err, data, res) {
                if (err) throw err;
                

                let order = JSON.parse(res);
                
  console.log('Reference number:', referenceNumber);
                if (order) {
                    let email = order.billing.email;
                  let billing_first_name = order.billing.first_name; // Extract the billing_first_name from the order object

                  let transporter = nodemailer.createTransport({
                    service: 'gmail',
                    auth: {
                      user: 'kunstinjekeuken@gmail.com',
                      pass: 'ndwu mhwn mjro fvyo'
                    }
                  });
                  let mailOptions = {
                    from: 'kunstinjekeuken@gmail.com',
                    to: email,
                    subject: 'Je bestelling is onderweg',
                    html: `
                      <p>Hoi ${billing_first_name},</p>
                      <p>Je bestelling is met de koerier mee. Zodra die vanavond in het DHL depot is aangekomen kun je hem volgen via deze Track&amp;Trace code: <a ${links[0]}</a>
                      </p>
                      <p>Controleer hem meteen nadat hij is afgeleverd. Graag hoor ik binnen twee dagen na ontvangst of hij in goede orde is aangekomen.</p>
                      <p>Ik wens je er alvast veel plezier mee! Je maakt mij heel blij met een voor en na foto シ</p>
                      <div style="color: grey;">
                      <p>
                        --<br>
                        <div style="line-height: 20px;">
                        Met vriendelijke groet,<br>
                        Leila<br>
                    <img src="https://i.imgur.com/jgVqcUZ.png" width="50" height="50" style="vertical-align: 0px;">
                        <p style="margin: 0;">Kunst in je keuken</p>
                       
                      </div>
                    
                    
                      <a href="http://www.kunstinjekeuken.nl">kunstinjekeuken.nl</a> | <a href="https://www.instagram.com/kunstinjekeuken/">Instagram</a> | <a href="https://www.facebook.com/leila.kunstinjekeuken/">Facebook</a></p>
                      
                    `
                  };
                  
                    
                      
                  
                  transporter.sendMail(mailOptions, (error, info) => {
                    if (error) {
                      console.log(error);
                    } else {
                      console.log('Email sent: ' + info.response);
                    }
                    });
                  }
                });
            });
          });
        });
        f.once('end', function() {
          imap.end();
        });
      });
    });
  });
  
  imap.once('error', function(err) {
    console.log(err);
  });
  
  imap.once('end', function() {
    console.log('Connection ended');
  });
  
  imap.connect();
  }
  cron.schedule('*/1 * * * *', checkEmails);