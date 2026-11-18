;;; E2P - Chuyen ELLIPSE / SPLINE thanh POLYLINE gom cac CUNG TRON (bulge), khong phai doan thang.
;;; Cach lam: tu dau duong cong, tim cung DAI NHAT (cung qua 3 diem dau - giua - cuoi) ma moi diem
;;; mau tren doan do lech <= sai so cho phep, roi di tiep tu cuoi cung do. => it dinh nhat co the:
;;; ellipse gan tron -> 1 cung, spline -> so cung toi thieu.
;;; Lenh: E2P  (doi tuong goc bi xoa, giu nguyen layer/mau/linetype; UNDO 1 lan la quay lai)
;; ponytail: cung 3 diem chi lien tuc vi tri (G0), o moi noi lech tiep tuyen rat nho (~ sai so);
;;           neu can tiep tuyen lien tuc tuyet doi (CNC chay muot) thi doi sang biarc.
(vl-load-com)

;; Diem tren duong cong tai chieu dai d, doi sang OCS, lay 2D.
;; Dung bien e, nrm, len, tol, h cua c:E2P (dynamic scope).
(defun _e2p-pt (d / p)
  (setq p (trans (cond ((vlax-curve-getPointAtDist e (min d len)))
                       ((vlax-curve-getEndPoint e)))
                 0 nrm))
  (list (car p) (cadr p)))

;; Sai lech cua q so voi duong tron qua p1 pm p2 (duong thang neu 3 diem thang hang).
(defun _e2p-dev (q p1 pm p2 / a b c)
  (setq a (mapcar '(lambda (x y) (/ (+ x y) 2.)) p1 pm)
        b (mapcar '(lambda (x y) (/ (+ x y) 2.)) pm p2)
        c (inters a (polar a (+ (angle p1 pm) (/ pi 2.)) 1.)
                  b (polar b (+ (angle pm p2) (/ pi 2.)) 1.) nil))
  (if c
    (abs (- (distance q c) (distance p1 c)))
    (/ (abs (- (* (- (car p2) (car p1)) (- (cadr q) (cadr p1)))
               (* (- (cadr p2) (cadr p1)) (- (car q) (car p1)))))
       (max (distance p1 p2) 1e-12))))

;; Cung qua 3 diem cua doan [d1 d2] co nam trong sai so tai moi diem mau khong?
;; Diem mau cach nhau <= h (it nhat 8 diem) de khong sot khuc uon cua spline.
(defun _e2p-ok (d1 d2 / p1 pm p2 n k ok)
  (setq p1 (_e2p-pt d1) pm (_e2p-pt (/ (+ d1 d2) 2.)) p2 (_e2p-pt d2)
        n  (max 8 (fix (/ (- d2 d1) h)))
        k  1
        ok T)
  (while (and ok (< k n))
    (setq ok (<= (_e2p-dev (_e2p-pt (+ d1 (* (- d2 d1) (/ k (float n))))) p1 pm p2) tol)
          k  (1+ k)))
  ok)

;; Bulge cua cung qua p1 pm p2: goc ngoat tai pm = 1/2 goc o tam -> bulge = tan(goc ngoat / 2)
(defun _e2p-bulge (p1 pm p2 / a)
  (setq a (- (angle pm p2) (angle p1 pm)))
  (cond ((> a pi) (setq a (- a pi pi)))
        ((< a (- pi)) (setq a (+ a pi pi))))
  (/ (sin (/ a 2.)) (max (cos (/ a 2.)) 1e-9)))

(defun c:E2P (/ ss i e ed nrm len cl h d lo hi m lst tol doc cnt)
  (setq *e2p-tol* (cond (*e2p-tol*) (0.01)))
  (initget 6)
  (setq tol (cond ((getdist (strcat "\nSai so cho phep <" (rtos *e2p-tol* 2 4) ">: ")))
                  (*e2p-tol*))
        *e2p-tol* tol)
  (if (setq ss (ssget "_:L" '((0 . "ELLIPSE,SPLINE"))))
    (progn
      (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)) cnt 0)
      (vla-StartUndoMark doc)
      (repeat (setq i (sslength ss))
        (setq e   (ssname ss (setq i (1- i)))
              ed  (entget e)
              nrm (cond ((cdr (assoc 210 ed))) ('(0. 0. 1.)))
              len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))
              cl  (vlax-curve-isClosed e)
              ;; buoc lay mau: spline nhieu diem dieu khien -> mau day hon
              h   (/ len (max 64 (* 8 (cond ((cdr (assoc 73 ed))) (0)))))
              d   0.
              lst nil)
        (while (< d (- len 1e-9))
          ;; duong kin: moi cung toi da nua chu vi (ellipse/tron kin -> 2 cung)
          (setq lo d
                hi (if cl (min len (+ d (/ len 2.))) len))
          (if (not (_e2p-ok d hi))
            (progn                              ; tim kiem nhi phan diem cuoi xa nhat con dat
              (while (> (- hi lo) (* len 1e-6))
                (setq m (/ (+ lo hi) 2.))
                (if (_e2p-ok d m) (setq lo m) (setq hi m)))
              (if (> lo d) (setq hi lo))))
          (setq lst (cons (list (_e2p-pt d)
                                (_e2p-bulge (_e2p-pt d) (_e2p-pt (/ (+ d hi) 2.)) (_e2p-pt hi)))
                          lst)
                d   hi))
        (or cl (setq lst (cons (list (_e2p-pt len) 0.) lst)))
        (setq lst (reverse lst))
        (if (entmakex
              (append
                (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity"))
                (vl-remove-if-not '(lambda (x) (member (car x) '(6 8 48 62 370 420 430))) ed)
                (list '(100 . "AcDbPolyline")
                      (cons 90 (length lst))
                      (cons 70 (if cl 1 0))
                      (cons 38 (caddr (trans (vlax-curve-getStartPoint e) 0 nrm))))
                (apply 'append (mapcar '(lambda (v) (list (cons 10 (car v)) (cons 42 (cadr v)))) lst))
                (list (cons 210 nrm))))
          (progn (entdel e) (setq cnt (1+ cnt)))))
      (vla-EndUndoMark doc)
      (princ (strcat "\nDa chuyen " (itoa cnt) " doi tuong thanh polyline cung tron."))))
  (princ))

(princ "\nE2P da nap. Go lenh E2P de chuyen Ellipse/Spline -> Polyline cung tron.")
(princ)
