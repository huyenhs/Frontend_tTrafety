#Region "IMPORTS"
Imports DAL.DBAccess
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlDbType
#End Region
Public Class front_end
#Region "DECLARE"

    'variable for connect database
    Private _obj_config As DAL.Config 'For connect

    Private _obj_db As DAL.DBAccess

    Private _str_cnn As String

    'attachment_type list
    Private _sql_home_page As String
    Private _sql_a_proc_new_breaker As String
    Private _sql_a_proc_select_all_danhmuc_by_id As String
    Private _sql_a_proc_select_all_tieu_diem_by_danhmucid As String
    Private _sql_a_check_danhmuc As String = "a_check_danhmuc"
    Private _sql_a_proc_select_all_danhmuc_with_danhmuc_cha_zero
    Private _sql_a_proc_select_detail_tintuc As String

    Private _sql_a_proc_vbqppl_select_all As String
    Private _sql_a_proc_vbqppl_side_bar As String
    Private _sql_a_proc_tinmoinhan As String
    Private _sql_a_proc_tinkhac As String
    Private _sql_a_check_insert_binh_luan As String
    Private _sql_a_proc_update_luot_xem As String
    Private _sql_a_proc_select_all_comment As String
    Private _sql_a_proc_select_vbqppl_tieude As String
    Private _sql_a_proc_select_vbqppl_types As String
    Private _sql_a_proc_vbqppl_select_all_search As String
    Private _sql_a_proc_vbqppl_detail As String
    Private _sql_a_proc_videos_select_all As String
    Private _sql_a_proc_select_videos_tieude As String
    Private _sql_a_proc_load_menu As String
    Private _sql_a_proc_load_tham_do_y_kien As String
    Private _sql_a_proc_insert_ketqua_thamdo_y_kien As String
    Private _sql_a_check_tham_do_y_kien As String
    Private _sql_a_proc_select_all_videos As String
    Private _sql_a_proc_select_view_videos As String
    Private _sql_a_proc_select_album_tieude As String
    Private _sql_a_proc_select_all_album As String
    Private _sql_a_proc_select_all_anh_in_album As String
    Private _sql_a_proc_select_danhsach_gop_y As String
    Private _sql_a_proc_select_hoidap_tieude As String
    Private _sql_a_proc_select_detail_gop_y As String = "a_proc_select_detail_gop_y"
    Private _sql_a_proc_them_gop_y As String = "a_proc_them_gop_y"
#End Region

#Region "CONSTRUCTOR"

    Public Sub New(ByVal Cnn As String)

        _obj_config = New DAL.Config

        _str_cnn = _obj_config.GetConnectionString(Cnn)

        _obj_db = New DAL.DBAccess

        initialization_variables()

    End Sub

    Private Sub initialization_variables()

        _sql_home_page = "a_proc_select_home"

        _sql_a_proc_new_breaker = "a_proc_new_breaker"
        _sql_a_proc_select_all_tieu_diem_by_danhmucid = "a_proc_select_all_tieu_diem_by_danhmucid"
        _sql_a_proc_select_all_danhmuc_by_id = "a_proc_select_all_danhmuc_by_id"
        _sql_a_proc_select_detail_tintuc = "a_proc_select_detail_tintuc"
        _sql_a_proc_select_all_danhmuc_with_danhmuc_cha_zero = "a_proc_select_all_danhmuc_with_danhmuc_cha_zero"

        _sql_a_proc_vbqppl_select_all = "a_proc_vbqppl_select_all"
        _sql_a_proc_vbqppl_side_bar = "a_proc_vbqppl_side_bar"

        _sql_a_proc_tinmoinhan = "a_proc_tinmoinhan"
        _sql_a_proc_tinkhac = "a_proc_tinkhac"

        _sql_a_proc_select_all_comment = "a_proc_select_all_comment"

        _sql_a_check_insert_binh_luan = "a_check_insert_binh_luan"
        _sql_a_proc_update_luot_xem = "a_proc_update_luot_xem"

        _sql_a_proc_select_vbqppl_tieude = "a_proc_select_vbqppl_tieude"

        _sql_a_proc_select_vbqppl_types = "a_proc_select_vbqppl_types"

        _sql_a_proc_vbqppl_select_all_search = "a_proc_vbqppl_select_all_search"

        _sql_a_proc_vbqppl_detail = "a_proc_vbqppl_detail"

        _sql_a_proc_videos_select_all = "a_proc_videos_select_all"
        _sql_a_proc_select_videos_tieude = "a_proc_select_videos_tieude"

        _sql_a_proc_load_menu = "a_proc_load_menu"

        _sql_a_proc_load_tham_do_y_kien = "a_proc_load_tham_do_y_kien"

        _sql_a_proc_insert_ketqua_thamdo_y_kien = "a_proc_insert_ketqua_thamdo_y_kien"

        _sql_a_check_tham_do_y_kien = "a_check_tham_do_y_kien"

        _sql_a_proc_select_all_videos = "a_proc_select_all_videos"

        _sql_a_proc_select_view_videos = "a_proc_select_view_videos"

        _sql_a_proc_select_album_tieude = "a_proc_select_album_tieude"

        _sql_a_proc_select_all_album = "a_proc_select_all_album"

        _sql_a_proc_select_all_anh_in_album = "a_proc_select_all_anh_in_album"

        _sql_a_proc_select_danhsach_gop_y = "a_proc_select_danhsach_gop_y"
        _sql_a_proc_select_hoidap_tieude = "a_proc_select_hoidap_tieude"

    End Sub


#End Region

    Public Function them_gop_y(ByVal HoTen As String, _
                                        ByVal Email As String, _
                                        ByVal DiaChi As String, _
                                        ByVal TieuDe As String, _
                                        ByVal NoiDungGopY As String) As Integer

        Dim sqlprm(4) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@HoTen"
        sqlprm(0).Value = HoTen

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@Email"
        sqlprm(1).Value = Email

        sqlprm(2) = New SqlParameter
        sqlprm(2).ParameterName = "@DiaChi"
        sqlprm(2).Value = DiaChi

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@TieuDe"
        sqlprm(3).Value = TieuDe

        sqlprm(4) = New SqlParameter
        sqlprm(4).ParameterName = "@NoiDungGopY"
        sqlprm(4).Value = NoiDungGopY


        Return _obj_db.SQlExecute(_str_cnn, _sql_a_proc_them_gop_y, sqlprm, True)
    End Function

    Public Function select_detail_gop_y(ByVal gopy_id As Integer) As DataTable

        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@gopy_id"
        sqlprm(0).Value = gopy_id

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_detail_gop_y, sqlprm, True)
    End Function


    Public Function select_hoidap_tieude() As DataTable


        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_hoidap_tieude, Nothing, False)
    End Function

    Public Function danh_sach_gop_y( _
       ByVal PageNum As Integer, _
ByVal PageSize As Integer, _
ByRef TotalRowsNum As Integer) As DataTable

        Dim sqlprm(2) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output


        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_select_danhsach_gop_y, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value
        Return tblb
    End Function


    Public Function select_anh_trong_album_theo_id(ByVal albumanhid As Integer) As DataTable

        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@albumid"
        sqlprm(0).Value = albumanhid

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_anh_in_album, sqlprm, True)
    End Function

    Public Function select_all_album() As DataTable


        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_album, Nothing, False)
    End Function

    Public Function select_album_tieu_de() As DataTable


        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_album_tieude, Nothing, False)
    End Function


    Public Function select_detailvideos(ByVal videoid As Integer) As DataSet
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@videoid"
        sqlprm(0).Value = videoid

        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_select_view_videos, sqlprm, True)
    End Function

    Public Function select_videos() As DataTable

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_videos, Nothing, False)
    End Function

    Public Function check_ket_qua(ByVal Ip As String) As DataTable
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@Ip"
        sqlprm(0).Value = Ip

        Return _obj_db.Filltable(_str_cnn, _sql_a_check_tham_do_y_kien, sqlprm, True)
    End Function

    Public Function insert_ketqua_thamdo_y_kien(ByVal Ip As String, ByVal ThamDoYKienId As Integer, _
                                                                ByVal TraLoi As Integer, _
                                                                ByVal HoTen As String, ByVal Email As String, _
                                                               ByVal SDT As String) As Integer


        Dim sqlprm(5) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@Ip"
        sqlprm(0).Value = Ip

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@ThamDoYKienId"
        sqlprm(1).Value = ThamDoYKienId

        sqlprm(2) = New SqlParameter
        sqlprm(2).ParameterName = "@TraLoi"
        sqlprm(2).Value = TraLoi

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@HoTen"
        sqlprm(3).Value = HoTen

        sqlprm(4) = New SqlParameter
        sqlprm(4).ParameterName = "@Email"
        sqlprm(4).Value = Email

        sqlprm(5) = New SqlParameter
        sqlprm(5).ParameterName = "@SDT"
        sqlprm(5).Value = SDT

        Return _obj_db.SQlExecute(_str_cnn, _sql_a_proc_insert_ketqua_thamdo_y_kien, sqlprm, True)

    End Function

    Public Function select_thamdoykien(ByVal thamdoid As Integer) As DataSet

        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@thamdoid"
        sqlprm(0).Value = thamdoid



        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_load_tham_do_y_kien, sqlprm, True)
    End Function

    Public Function load_menu() As DataTable

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_load_menu, Nothing, False)
    End Function

    Public Function select_videos_tieude() As DataTable

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_videos_tieude, Nothing, False)
    End Function

    'a_proc_select_vbqppl_types

    Public Function select_vbqppl_types() As DataSet

        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_select_vbqppl_types, Nothing, False)
    End Function

    Public Function select_vbqppl_chitiet(ByVal VanBanID As Integer, ByVal counts As Integer) As DataSet

        Dim sqlprm(1) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@VanBanID"
        sqlprm(0).Value = VanBanID

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@counts"
        sqlprm(1).Value = counts

        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_vbqppl_detail, sqlprm, True)
    End Function

    Public Function select_vbqppl_tieude() As DataTable

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_vbqppl_tieude, Nothing, False)
    End Function

    Public Function chitiet(ByVal tinid As Integer) As DataTable

        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@tintucid"
        sqlprm(0).Value = tinid

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_detail_tintuc, sqlprm, True)

    End Function

    Public Function new_breaker() As DataTable

        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_new_breaker, Nothing, False)

    End Function

    Public Function front_end_homepage() As DataSet

        Return _obj_db.FillDataSet(_str_cnn, _sql_home_page, Nothing, False)

    End Function

    Public Function front_end_select_articles_by_danhmuc( _
        ByVal PageNum As Integer, _
 ByVal PageSize As Integer, _
 ByRef TotalRowsNum As Integer, ByVal danhmucid As Integer) As DataTable

        Dim sqlprm(3) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@danhmucid"
        sqlprm(3).Value = danhmucid

        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_danhmuc_by_id, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value
        Return tblb
    End Function

    Public Function front_end_select_articles_by_danhmuc_tieudiem(ByVal danhmucid As Integer) As DataTable
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@danhmucid"
        sqlprm(0).Value = danhmucid
        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_tieu_diem_by_danhmucid, sqlprm, True)

    End Function

    '_sql_a_check_danhmuc

    Public Function front_end_check_danhmuc(ByVal danhmucid As Integer) As DataTable
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@danhmucid"
        sqlprm(0).Value = danhmucid
        Return _obj_db.Filltable(_str_cnn, _sql_a_check_danhmuc, sqlprm, True)

    End Function

    '_sql_a_proc_select_all_danhmuc_with_danhmuc_cha_zero
    Public Function front_end_danhmuc_with_parent_zero(ByVal danhmucid As Integer) As DataSet
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@danhmucid"
        sqlprm(0).Value = danhmucid
        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_select_all_danhmuc_with_danhmuc_cha_zero, sqlprm, True)

    End Function

    Public Function vbqppl_side_bar() As DataSet
        Return _obj_db.FillDataSet(_str_cnn, _sql_a_proc_vbqppl_side_bar, Nothing, False)

    End Function

    Public Function vbqppl_select_all( _
        ByVal PageNum As Integer, _
 ByVal PageSize As Integer, _
 ByRef TotalRowsNum As Integer) As DataTable

        Dim sqlprm(2) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output



        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_vbqppl_select_all, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value
        Return tblb
    End Function


    Public Function vbqppl_search( _
        ByVal PageNum As Integer, _
 ByVal PageSize As Integer, _
 ByRef TotalRowsNum As Integer, _
 ByVal MaLinhVuc As Integer, ByVal MaHinhThuc As Integer, ByVal MaCQBH As Integer, ByVal MaChuyenNganh As Integer, _
 ByVal keyword As String, ByVal usekeyword As Boolean) As DataTable

        Dim sqlprm(8) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@MaLinhVuc"
        sqlprm(3).Value = MaLinhVuc

        sqlprm(4) = New SqlParameter
        sqlprm(4).ParameterName = "@MaHinhThuc"
        sqlprm(4).Value = MaHinhThuc

        sqlprm(5) = New SqlParameter
        sqlprm(5).ParameterName = "@MaCQBH"
        sqlprm(5).Value = MaCQBH

        sqlprm(6) = New SqlParameter
        sqlprm(6).ParameterName = "@MaChuyenNganh"
        sqlprm(6).Value = MaChuyenNganh

        sqlprm(7) = New SqlParameter
        sqlprm(7).ParameterName = "@keyword"
        sqlprm(7).Value = keyword

        sqlprm(8) = New SqlParameter
        sqlprm(8).ParameterName = "@usekeyword"
        sqlprm(8).Value = usekeyword



        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_vbqppl_select_all_search, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value
        Return tblb
    End Function

    '
    Public Function front_end_tinmoinhan(ByVal tintucid As Integer) As DataTable
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@tintucid"
        sqlprm(0).Value = tintucid
        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_tinmoinhan, sqlprm, True)

    End Function

    Public Function front_end_tintuckhacs(ByVal tintucid As Integer) As DataTable
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@tintucid"
        sqlprm(0).Value = tintucid
        Return _obj_db.Filltable(_str_cnn, _sql_a_proc_tinkhac, sqlprm, True)

    End Function
    '_sql_a_check_insert_binh_luan

    Public Function insert_binhluan(ByVal HoTen As String, _
                                         ByVal Email As String, _
                                         ByVal NoiDung As String, ByVal BinhLuanChoTinTuc As Integer) As Integer

        Dim sqlprm(3) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@HoTen"
        sqlprm(0).Value = HoTen

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@Email"
        sqlprm(1).Value = Email

        sqlprm(2) = New SqlParameter
        sqlprm(2).ParameterName = "@NoiDung"
        sqlprm(2).Value = NoiDung

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@BinhLuanChoTinTuc"
        sqlprm(3).Value = BinhLuanChoTinTuc


        Return _obj_db.SQlExecute(_str_cnn, _sql_a_check_insert_binh_luan, sqlprm, True)

    End Function

    Public Function cap_nhat_luot_xem(ByVal tintucid As Integer) As Integer
        Dim sqlprm(0) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@tinid"
        sqlprm(0).Value = tintucid
        Return _obj_db.SQlExecute(_str_cnn, _sql_a_proc_update_luot_xem, sqlprm, True)

    End Function

    '_sql_a_proc_select_all_comment
    Public Function select_all_comment(ByVal tintucid As Integer, ByVal PageNum As Integer, _
 ByVal PageSize As Integer, _
 ByRef TotalRowsNum As Integer) As DataTable
        Dim sqlprm(3) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output

        sqlprm(3) = New SqlParameter
        sqlprm(3).ParameterName = "@tintucid"
        sqlprm(3).Value = tintucid

        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_select_all_comment, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value


        Return tblb

    End Function

    '_sql_a_proc_videos_select_all
    Public Function video_select_all(ByVal PageNum As Integer, _
 ByVal PageSize As Integer, _
 ByRef TotalRowsNum As Integer) As DataTable


        Dim sqlprm(2) As SqlParameter

        sqlprm(0) = New SqlParameter
        sqlprm(0).ParameterName = "@PageNum"
        sqlprm(0).Value = PageNum

        sqlprm(1) = New SqlParameter
        sqlprm(1).ParameterName = "@PageSize"
        sqlprm(1).Value = PageSize

        sqlprm(2) = New SqlParameter("@TotalRowsNum", SqlDbType.Int)
        sqlprm(2).Direction = ParameterDirection.Output



        Dim tblb As New DataTable
        tblb = _obj_db.Filltable(_str_cnn, _sql_a_proc_videos_select_all, sqlprm, True)
        TotalRowsNum = sqlprm(2).Value
        Return tblb
    End Function
End Class
