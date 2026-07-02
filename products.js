const XLSX = require('xlsx');

const products = [
  // TEL
  {name:'Sarso Tel', category:'Tel', s1:'500ml', p1:85, s2:'1L', p2:160, s3:'2L', p3:310},
  {name:'Sunflower Tel', category:'Tel', s1:'500ml', p1:90, s2:'1L', p2:175, s3:'2L', p3:340},
  {name:'Mustard Tel', category:'Tel', s1:'500ml', p1:80, s2:'1L', p2:155, s3:'2L', p3:300},
  {name:'Him Gange Tel', category:'Tel', s1:'100ml', p1:55, s2:'200ml', p2:105, s3:'500ml', p3:250},
  {name:'Coconut Tel', category:'Tel', s1:'100ml', p1:55, s2:'200ml', p2:105, s3:'500ml', p3:250},
  // SHAMPOO
  {name:'Head & Shoulders', category:'Shampoo', s1:'100ml', p1:99, s2:'200ml', p2:185, s3:'400ml', p3:350},
  {name:'Pantene Shampoo', category:'Shampoo', s1:'100ml', p1:89, s2:'200ml', p2:165, s3:'400ml', p3:320},
  {name:'Dove Shampoo', category:'Shampoo', s1:'100ml', p1:95, s2:'200ml', p2:180, s3:'400ml', p3:340},
  {name:'Clinic Plus', category:'Shampoo', s1:'100ml', p1:45, s2:'200ml', p2:85, s3:'400ml', p3:160},
  {name:'Sunsilk Shampoo', category:'Shampoo', s1:'100ml', p1:55, s2:'200ml', p2:99, s3:'400ml', p3:185},
  {name:'Himalaya Shampoo', category:'Shampoo', s1:'100ml', p1:75, s2:'200ml', p2:140, s3:'400ml', p3:270},
  {name:'Patanjali Kesh Kanti', category:'Shampoo', s1:'100ml', p1:60, s2:'200ml', p2:115, s3:'400ml', p3:220},
  // GLUCON D
  {name:'Glucon D Orange', category:'Health Drink', s1:'100g', p1:45, s2:'200g', p2:85, s3:'500g', p3:195},
  {name:'Glucon D Lemon', category:'Health Drink', s1:'100g', p1:45, s2:'200g', p2:85, s3:'500g', p3:195},
  {name:'Glucon D Mango', category:'Health Drink', s1:'100g', p1:45, s2:'200g', p2:85, s3:'500g', p3:195},
  {name:'Horlicks', category:'Health Drink', s1:'200g', p1:120, s2:'500g', p2:280, s3:'1kg', p3:540},
  {name:'Boost', category:'Health Drink', s1:'200g', p1:130, s2:'500g', p2:300, s3:'1kg', p3:580},
  {name:'Bournvita', category:'Health Drink', s1:'200g', p1:125, s2:'500g', p2:290, s3:'1kg', p3:560},
  // MEDICINE
  {name:'Paracetamol', category:'Medicine', s1:'10 tab', p1:12, s2:'15 tab', p2:18, s3:'30 tab', p3:35},
  {name:'Disprin', category:'Medicine', s1:'10 tab', p1:15, s2:'15 tab', p2:22, s3:'30 tab', p3:42},
  {name:'Crocin', category:'Medicine', s1:'10 tab', p1:18, s2:'15 tab', p2:26, s3:'30 tab', p3:50},
  // SOAP
  {name:'Dettol Soap', category:'Soap', s1:'75g', p1:35, s2:'125g', p2:55, s3:'250g', p3:105},
  {name:'Lux Soap', category:'Soap', s1:'75g', p1:30, s2:'125g', p2:48, s3:'250g', p3:92},
  {name:'Lifebuoy Soap', category:'Soap', s1:'75g', p1:28, s2:'125g', p2:44, s3:'250g', p3:85},
  {name:'Dove Soap', category:'Soap', s1:'75g', p1:55, s2:'125g', p2:88, s3:'250g', p3:168},
  {name:'Him Gange Soap', category:'Soap', s1:'75g', p1:25, s2:'125g', p2:40, s3:'250g', p3:75},
  {name:'Pears Soap', category:'Soap', s1:'75g', p1:42, s2:'125g', p2:68, s3:'250g', p3:130},
  // TOOTHPASTE
  {name:'Colgate', category:'Toothpaste', s1:'50g', p1:35, s2:'100g', p2:65, s3:'200g', p3:125},
  {name:'Pepsodent', category:'Toothpaste', s1:'50g', p1:30, s2:'100g', p2:58, s3:'200g', p3:112},
  {name:'Dabur Red', category:'Toothpaste', s1:'50g', p1:38, s2:'100g', p2:72, s3:'200g', p3:140},
  {name:'Sensodyne', category:'Toothpaste', s1:'50g', p1:85, s2:'100g', p2:160, s3:'200g', p3:310},
  {name:'Patanjali Dant Kanti', category:'Toothpaste', s1:'50g', p1:35, s2:'100g', p2:65, s3:'200g', p3:125},
  // HANDWASH
  {name:'Dettol Handwash', category:'Handwash', s1:'200ml', p1:65, s2:'500ml', p2:145, s3:'1L', p3:280},
  {name:'Lifebuoy Handwash', category:'Handwash', s1:'200ml', p1:55, s2:'500ml', p2:125, s3:'1L', p3:240},
  {name:'Savlon Handwash', category:'Handwash', s1:'200ml', p1:60, s2:'500ml', p2:135, s3:'1L', p3:260},
  // DETERGENT
  {name:'Surf Excel', category:'Detergent', s1:'200g', p1:35, s2:'500g', p2:80, s3:'1kg', p3:155},
  {name:'Ariel', category:'Detergent', s1:'200g', p1:40, s2:'500g', p2:92, s3:'1kg', p3:178},
  {name:'Tide', category:'Detergent', s1:'200g', p1:30, s2:'500g', p2:68, s3:'1kg', p3:130},
  {name:'Rin', category:'Detergent', s1:'200g', p1:22, s2:'500g', p2:50, s3:'1kg', p3:95},
  {name:'Nirma', category:'Detergent', s1:'200g', p1:15, s2:'500g', p2:35, s3:'1kg', p3:68},
  {name:'Wheel', category:'Detergent', s1:'200g', p1:18, s2:'500g', p2:42, s3:'1kg', p3:80},
  // DISHWASH
  {name:'Vim', category:'Dishwash', s1:'200g', p1:30, s2:'500g', p2:68, s3:'1kg', p3:130},
  {name:'Pril', category:'Dishwash', s1:'200ml', p1:35, s2:'500ml', p2:80, s3:'1L', p3:155},
  {name:'Exo', category:'Dishwash', s1:'200g', p1:25, s2:'500g', p2:58, s3:'1kg', p3:110},
  // HAIR OIL
  {name:'Dabur Amla Oil', category:'Hair Oil', s1:'100ml', p1:60, s2:'200ml', p2:115, s3:'500ml', p3:270},
  {name:'Bajaj Almond Oil', category:'Hair Oil', s1:'100ml', p1:65, s2:'200ml', p2:125, s3:'500ml', p3:300},
  {name:'Vatika Hair Oil', category:'Hair Oil', s1:'100ml', p1:70, s2:'200ml', p2:135, s3:'500ml', p3:320},
  {name:'Nihar Coconut Oil', category:'Hair Oil', s1:'100ml', p1:50, s2:'200ml', p2:95, s3:'500ml', p3:225},
  {name:'Parachute Oil', category:'Hair Oil', s1:'100ml', p1:55, s2:'200ml', p2:105, s3:'500ml', p3:250},
  // CREAM/LOTION
  {name:'Vaseline Lotion', category:'Cream', s1:'100ml', p1:75, s2:'200ml', p2:140, s3:'400ml', p3:270},
  {name:'Nivea Lotion', category:'Cream', s1:'100ml', p1:85, s2:'200ml', p2:160, s3:'400ml', p3:310},
  {name:'Ponds Cream', category:'Cream', s1:'50g', p1:55, s2:'100g', p2:105, s3:'200g', p3:200},
  {name:'Boroplus Cream', category:'Cream', s1:'50g', p1:45, s2:'100g', p2:85, s3:'200g', p3:160},
  {name:'Fair & Lovely', category:'Cream', s1:'50g', p1:65, s2:'100g', p2:125, s3:'200g', p3:240},
  // FACE WASH
  {name:'Himalaya Face Wash', category:'Face Wash', s1:'50ml', p1:65, s2:'100ml', p2:120, s3:'200ml', p3:230},
  {name:'Garnier Face Wash', category:'Face Wash', s1:'50ml', p1:75, s2:'100ml', p2:140, s3:'200ml', p3:270},
  {name:'Ponds Face Wash', category:'Face Wash', s1:'50ml', p1:70, s2:'100ml', p2:130, s3:'200ml', p3:250},
  // FLOOR/TOILET CLEANER
  {name:'Harpic', category:'Cleaner', s1:'200ml', p1:55, s2:'500ml', p2:125, s3:'1L', p3:240},
  {name:'Domex', category:'Cleaner', s1:'200ml', p1:50, s2:'500ml', p2:115, s3:'1L', p3:220},
  {name:'Lizol', category:'Cleaner', s1:'200ml', p1:60, s2:'500ml', p2:135, s3:'1L', p3:260},
  {name:'Colin', category:'Cleaner', s1:'200ml', p1:65, s2:'500ml', p2:148, s3:'1L', p3:285},
  {name:'Phenyl', category:'Cleaner', s1:'200ml', p1:25, s2:'500ml', p2:55, s3:'1L', p3:105},
  // MOSQUITO
  {name:'Good Knight', category:'Mosquito', s1:'45 nights', p1:55, s2:'90 nights', p2:105, s3:'180 nights', p3:200},
  {name:'Mortein', category:'Mosquito', s1:'45 nights', p1:60, s2:'90 nights', p2:115, s3:'180 nights', p3:220},
  {name:'Hit Spray', category:'Mosquito', s1:'200ml', p1:85, s2:'400ml', p2:160, s3:'625ml', p3:245},
  {name:'Baygon Spray', category:'Mosquito', s1:'200ml', p1:90, s2:'400ml', p2:170, s3:'625ml', p3:260},
  // COFFEE/TEA
  {name:'Nescafe', category:'Coffee', s1:'50g', p1:120, s2:'100g', p2:230, s3:'200g', p3:450},
  {name:'Bru Coffee', category:'Coffee', s1:'50g', p1:95, s2:'100g', p2:180, s3:'200g', p3:350},
  // ATTA
  {name:'Aashirvaad Atta', category:'Atta', s1:'1kg', p1:50, s2:'5kg', p2:235, s3:'10kg', p3:460},
  {name:'Annapurna Atta', category:'Atta', s1:'1kg', p1:45, s2:'5kg', p2:210, s3:'10kg', p3:410},
  {name:'Pillsbury Atta', category:'Atta', s1:'1kg', p1:48, s2:'5kg', p2:225, s3:'10kg', p3:440},
  {name:'Fortune Atta', category:'Atta', s1:'1kg', p1:46, s2:'5kg', p2:215, s3:'10kg', p3:420},
  // RICE
  {name:'India Gate Rice', category:'Rice', s1:'1kg', p1:95, s2:'5kg', p2:450, s3:'10kg', p3:880},
  {name:'Daawat Rice', category:'Rice', s1:'1kg', p1:90, s2:'5kg', p2:425, s3:'10kg', p3:830},
  {name:'Long Mogra Rice', category:'Rice', s1:'1kg', p1:85, s2:'5kg', p2:400, s3:'10kg', p3:780},
  // DAL
  {name:'Toor Dal', category:'Dal', s1:'500g', p1:55, s2:'1kg', p2:105, s3:'5kg', p3:510},
  {name:'Moong Dal', category:'Dal', s1:'500g', p1:60, s2:'1kg', p2:115, s3:'5kg', p3:560},
  {name:'Chana Dal', category:'Dal', s1:'500g', p1:50, s2:'1kg', p2:95, s3:'5kg', p3:460},
  {name:'Masoor Dal', category:'Dal', s1:'500g', p1:48, s2:'1kg', p2:92, s3:'5kg', p3:445},
  // MASALE
  {name:'MDH Garam Masala', category:'Masala', s1:'50g', p1:35, s2:'100g', p2:65, s3:'500g', p3:300},
  {name:'Everest Masala', category:'Masala', s1:'50g', p1:32, s2:'100g', p2:60, s3:'500g', p3:280},
  {name:'Catch Masala', category:'Masala', s1:'50g', p1:30, s2:'100g', p2:55, s3:'500g', p3:260},
  // NOODLES
  {name:'Maggi Noodles', category:'Noodles', s1:'70g', p1:14, s2:'140g', p2:26, s3:'420g', p3:75},
  {name:'Yippee Noodles', category:'Noodles', s1:'70g', p1:12, s2:'140g', p2:22, s3:'420g', p3:65},
  {name:'Top Ramen', category:'Noodles', s1:'70g', p1:12, s2:'140g', p2:22, s3:'420g', p3:65},
  // BISCUIT
  {name:'Parle G', category:'Biscuit', s1:'100g', p1:10, s2:'200g', p2:18, s3:'500g', p3:45},
  {name:'Britannia', category:'Biscuit', s1:'100g', p1:20, s2:'200g', p2:38, s3:'500g', p3:92},
  // SALT
  {name:'Tata Salt', category:'Salt', s1:'500g', p1:12, s2:'1kg', p2:22, s3:'2kg', p3:42},
  {name:'Catch Salt', category:'Salt', s1:'500g', p1:14, s2:'1kg', p2:26, s3:'2kg', p3:50},
];

// Har variant ka starting stock. Isse badalne ke liye yahan number change karo
// aur `node products.js` dobara chalao (isse orders.xlsx/live stock reset nahi hota,
// sirf products.xlsx re-generate hota hai).
const STARTING_STOCK = 100;

const rows = [];
let sr = 1;
products.forEach(p => {
  rows.push({ 'Sr': sr++, 'Product': p.name, 'Category': p.category, 'Size': p.s1, 'Price (Rs)': p.p1, 'Stock': STARTING_STOCK });
  rows.push({ 'Sr': sr++, 'Product': p.name, 'Category': p.category, 'Size': p.s2, 'Price (Rs)': p.p2, 'Stock': STARTING_STOCK });
  rows.push({ 'Sr': sr++, 'Product': p.name, 'Category': p.category, 'Size': p.s3, 'Price (Rs)': p.p3, 'Stock': STARTING_STOCK });
});

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet(rows);
ws['!cols'] = [{wch:5},{wch:25},{wch:15},{wch:10},{wch:12},{wch:10}];
XLSX.utils.book_append_sheet(wb, ws, 'Products');
XLSX.writeFile(wb, 'products.xlsx');
console.log('Done! ' + rows.length + ' products saved, har variant ka stock ' + STARTING_STOCK + '!');