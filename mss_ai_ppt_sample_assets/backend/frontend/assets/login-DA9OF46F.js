import{c as d,r as t,j as e,A as h,a as p,R as j}from"./base-DPddANql.js";/**
 * @license lucide-react v0.546.0 - ISC
 *
 * This source code is licensed under the ISC license.
 * See the LICENSE file in the root directory of this source tree.
 */const g=[["rect",{width:"18",height:"11",x:"3",y:"11",rx:"2",ry:"2",key:"1w4ew1"}],["path",{d:"M7 11V7a5 5 0 0 1 10 0v4",key:"fwvmzm"}]],y=d("lock",g);/**
 * @license lucide-react v0.546.0 - ISC
 *
 * This source code is licensed under the ISC license.
 * See the LICENSE file in the root directory of this source tree.
 */const v=[["path",{d:"M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2",key:"975kel"}],["circle",{cx:"12",cy:"7",r:"4",key:"17ys0d"}]],f=d("user",v);function w(){const[r,m]=t.useState(""),[a,u]=t.useState(""),[o,i]=t.useState(!1),[c,n]=t.useState(""),x=async s=>{if(s.preventDefault(),!r.trim()||!a){n("请输入用户名和密码");return}i(!0),n("");try{await h.login(r.trim(),a),window.location.href="/ui/admin.html"}catch(l){n(l instanceof Error?l.message:"登录失败")}finally{i(!1)}};return e.jsx("div",{className:"login-page",children:e.jsxs("div",{className:"login-card",children:[e.jsx("div",{className:"login-logo",children:"M"}),e.jsx("h1",{children:"MSS AI PPT"}),e.jsx("p",{children:"管理后台登录"}),e.jsxs("form",{onSubmit:x,className:"login-form",children:[e.jsxs("label",{children:[e.jsxs("span",{children:[e.jsx(f,{size:14})," 用户名"]}),e.jsx("input",{value:r,onChange:s=>m(s.target.value),autoComplete:"username"})]}),e.jsxs("label",{children:[e.jsxs("span",{children:[e.jsx(y,{size:14})," 密码"]}),e.jsx("input",{type:"password",value:a,onChange:s=>u(s.target.value),autoComplete:"current-password"})]}),c&&e.jsx("div",{className:"login-error",children:c}),e.jsx("button",{type:"submit",className:"btn-primary",disabled:o,children:o?"登录中...":"登录"})]})]})})}p.createRoot(document.getElementById("root")).render(e.jsx(j.StrictMode,{children:e.jsx(w,{})}));
