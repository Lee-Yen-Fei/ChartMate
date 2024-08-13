import express from 'express';
import cors from 'cors';

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

// Example route
app.get('/', (req, res) => {
  res.send('Backend for Canva App is running!');
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
