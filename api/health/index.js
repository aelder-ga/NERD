module.exports = async function (context) {
  context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body: { status: 'ok', service: 'NERD API' } };
};
