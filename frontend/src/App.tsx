import "./index.css";
import { Link, Navigate, Route, Routes, useLocation } from "react-router-dom";
import { CreatePackage } from "./pages/CreatePackage";
import { ClientsList } from "./pages/ClientsList";
import { ClientDetails } from "./pages/ClientDetails";

function App() {
  const loc = useLocation();
  return (
    <div>
      <div className="container" style={{ paddingBottom: 0 }}>
        <div className="row" style={{ justifyContent: "space-between" }}>
          <div className="row">
            <Link className="link" to="/create" style={{ fontWeight: 700 }}>
              Visa
            </Link>
            <Link className="link" to="/create" style={{ fontWeight: loc.pathname.startsWith("/create") ? 700 : 600 }}>
              Create package
            </Link>
            <Link className="link" to="/clients" style={{ fontWeight: loc.pathname.startsWith("/clients") ? 700 : 600 }}>
              Clients
            </Link>
          </div>
          <div className="hint">MVP</div>
        </div>
      </div>

      <Routes>
        <Route path="/" element={<Navigate to="/create" replace />} />
        <Route path="/create" element={<CreatePackage />} />
        <Route path="/clients" element={<ClientsList />} />
        <Route path="/clients/:clientId" element={<ClientDetails />} />
        <Route path="*" element={<Navigate to="/create" replace />} />
      </Routes>
    </div>
  );
}

export default App
