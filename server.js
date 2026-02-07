const { app, PORT } = require('./backendApp');

app.listen(PORT, () => {
  console.log(`Quiz app running at http://localhost:${PORT}`);
});
